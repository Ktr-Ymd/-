"""明細書チェックくん
==========================

特許明細書案（.docx）を読み込み、以下の観点でチェックを行います。

* 審査基準や特許法上重要となる基本セクションの有無
* 実施可能要件・サポート要件を満たすための最低限の記載が揃っているか
* 形式面（請求項番号の連番、和文中の半角記号など）の揺れ
* プレースホルダーや誤記の疑いがある文字列

指摘事項ごとに内容を表示し、自動修正が可能な項目については
コマンドライン上で修正を適用して新しいWordファイルを出力できます。

使い方の概要
---------------

.. code-block:: bash

    $ python 明細書チェックくん.py <入力ファイル.docx>

チェック完了後、画面の案内に従って自動修正の適用有無やメモを入力すると、
レポート（Markdown形式）と必要に応じて修正版の明細書が生成されます。
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import re
import sys
import textwrap
from dataclasses import dataclass, field
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZipInfo


# === 基本的な docx ハンドラ ==================================================


_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


class DocxDocument:
    """非常に小さな docx ラッパー。

    python-docx 等の外部ライブラリを利用できない環境でも、
    word/document.xml の読み込みと単純な文字列置換による更新ができるようにする。
    """

    def __init__(self, path: str) -> None:
        if not os.path.exists(path):
            raise FileNotFoundError(path)
        self.path = os.path.abspath(path)
        self._zip_infos: Dict[str, ZipInfo] = {}
        self._zip_content: Dict[str, bytes] = {}
        with ZipFile(self.path, "r") as zf:
            for info in zf.infolist():
                self._zip_infos[info.filename] = info
                self._zip_content[info.filename] = zf.read(info.filename)
        try:
            self._document_xml = self._zip_content["word/document.xml"]
        except KeyError as exc:  # pragma: no cover - docx に必須のファイル
            raise ValueError("word/document.xml が見つかりません。破損した docx です。") from exc
        self._tree = ET.fromstring(self._document_xml)
        self._paragraph_cache: Optional[List[str]] = None

    # --- 読み取り系 ---------------------------------------------------------
    @property
    def paragraphs(self) -> List[str]:
        if self._paragraph_cache is None:
            paras: List[str] = []
            for p in self._tree.findall(".//w:p", namespaces=_NS):
                texts = [t.text for t in p.findall(".//w:t", namespaces=_NS) if t.text]
                if texts:
                    paras.append("".join(texts))
            self._paragraph_cache = paras
        return list(self._paragraph_cache)

    @property
    def full_text(self) -> str:
        return "\n".join(self.paragraphs)

    # --- 書き換え系 ---------------------------------------------------------
    def replace_text(self, old: str, new: str) -> bool:
        """w:t 要素単位で単純な置換を実施する。"""

        if not old:
            return False
        changed = False
        for t in self._tree.findall(".//w:t", namespaces=_NS):
            if t.text and old in t.text:
                t.text = t.text.replace(old, new)
                changed = True
        if changed:
            self._paragraph_cache = None
        return changed

    def save(self, output_path: str) -> None:
        xml_bytes = ET.tostring(
            self._tree, encoding="utf-8", xml_declaration=True
        )
        self._zip_content["word/document.xml"] = xml_bytes
        output_path = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        with ZipFile(output_path, "w") as zf:
            for filename, data in self._zip_content.items():
                info = self._zip_infos.get(filename)
                if info is not None:
                    # ZipInfo を複製しないと書き込み時に属性を共有してしまうためコピー
                    new_info = ZipInfo(filename)
                    new_info.date_time = info.date_time
                    new_info.compress_type = info.compress_type
                    new_info.comment = info.comment
                    new_info.extra = info.extra
                    zf.writestr(new_info, data)
                else:
                    zf.writestr(filename, data)


# === チェック結果を表現するデータモデル =====================================


Severity = str


@dataclass
class AutoFix:
    """自動修正に必要な情報。"""

    label: str
    apply: Callable[[DocxDocument], bool]


@dataclass
class Finding:
    identifier: str
    title: str
    message: str
    category: str
    severity: Severity
    location: str
    context: str = ""
    suggestion: str = ""
    autofix: Optional[AutoFix] = None


@dataclass
class FindingResolution:
    finding: Finding
    status: str
    note: str = ""
    applied_fix: bool = False


# === ユーティリティ ===========================================================


_FULLWIDTH_DIGITS = str.maketrans({
    "０": "0",
    "１": "1",
    "２": "2",
    "３": "3",
    "４": "4",
    "５": "5",
    "６": "6",
    "７": "7",
    "８": "8",
    "９": "9",
})


def _normalize_digit_string(text: str) -> str:
    return text.translate(_FULLWIDTH_DIGITS)


def _normalize_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip())


def _extract_claims(paragraphs: Sequence[str]) -> List[Tuple[int, int, str]]:
    pattern = re.compile(r"(?:【)?請求項[ 　]*([0-9０-９]+)")
    claims: List[Tuple[int, int, str]] = []
    for idx, para in enumerate(paragraphs):
        if "請求項" not in para:
            continue
        match = pattern.search(para)
        if not match:
            continue
        number = match.group(1)
        try:
            num_value = int(_normalize_digit_string(number))
        except ValueError:
            continue
        claim_body = para[match.end() :].strip()
        claims.append((num_value, idx, claim_body))
    return claims


def _extract_keywords(text: str) -> List[str]:
    candidates = set()
    for word in re.findall(r"[A-Za-z0-9\-]{3,}", text):
        candidates.add(word.lower())
    for word in re.findall(r"[一-龥]{2,}", text):
        if len(word) >= 2:
            candidates.add(word)
    for quoted in re.findall(r"「([^」]{2,})」", text):
        cleaned = quoted.strip()
        if len(cleaned) >= 2:
            candidates.add(cleaned)
    return sorted(candidates)


def _paragraph_snippet(paragraphs: Sequence[str], index: int) -> str:
    if not 0 <= index < len(paragraphs):
        return ""
    snippet = _normalize_whitespace(paragraphs[index])
    return snippet[:120] + ("…" if len(snippet) > 120 else "")


# === チェックロジック =========================================================


def check_required_sections(doc: DocxDocument) -> List[Finding]:
    paragraphs = doc.paragraphs
    joined = doc.full_text
    requirements = [
        (
            "REQ-001",
            "必須セクション",
            "法的要件",
            "高",
            "『特許請求の範囲』が見当たりません。特許法第36条5項に従い必須です。",
            "全体",
            "特許請求の範囲の見出しと請求項群を追加してください。",
            "特許請求の範囲",
        ),
        (
            "REQ-002",
            "必須セクション",
            "法的要件",
            "高",
            "『発明の詳細な説明』が不足しています。特許法第36条4項1号の実施可能要件に抵触する恐れがあります。",
            "全体",
            "発明の詳細な説明の見出しと、少なくとも実施の形態の説明を記載してください。",
            "発明の詳細な説明",
        ),
        (
            "REQ-003",
            "必須セクション",
            "法的要件",
            "中",
            "『課題を解決するための手段』の記載が確認できません。審査基準上、課題と解決手段を明確にすることが推奨されます。",
            "全体",
            "発明が解決しようとする課題に対応する手段を明記してください。",
            "課題を解決するための手段",
        ),
        (
            "REQ-004",
            "必須セクション",
            "法的要件",
            "中",
            "『図面の簡単な説明』が不足しています。図面がある場合は審査基準上求められます。",
            "全体",
            "図面の簡単な説明の見出しと各図の概要を記載してください。",
            "図面の簡単な説明",
        ),
    ]

    findings: List[Finding] = []
    for identifier, title, category, severity, message, location, suggestion, keyword in requirements:
        if keyword not in joined:
            findings.append(
                Finding(
                    identifier=identifier,
                    title=title,
                    message=message,
                    category=category,
                    severity=severity,
                    location=location,
                    suggestion=suggestion,
                )
            )

    # 図面の簡単な説明は図の記載が無い場合は指摘不要
    has_figure_ref = any(re.search(r"図[0-9０-９]", p) for p in paragraphs)
    if not has_figure_ref:
        findings = [f for f in findings if f.identifier != "REQ-004"]
    return findings


def check_enablement(doc: DocxDocument) -> List[Finding]:
    paragraphs = doc.paragraphs
    joined = doc.full_text
    findings: List[Finding] = []
    has_example = any("実施例" in p or "実施の形態" in p for p in paragraphs)
    if not has_example:
        findings.append(
            Finding(
                identifier="ENB-001",
                title="実施可能要件の検討",
                message="実施例または実施の形態の記載が見当たりません。特許法36条4項1号に基づき、実施できる程度に発明を開示する必要があります。",
                category="実施可能要件",
                severity="高",
                location="発明の詳細な説明",
                suggestion="具体的な実施例（構成、手順、条件等）を記載してください。",
            )
        )
    keywords = ["効果", "発明の効果"]
    if not any(keyword in joined for keyword in keywords):
        findings.append(
            Finding(
                identifier="ENB-002",
                title="効果記載の不足",
                message="発明の効果に関する見出しまたは記載が不足しています。課題と効果の対応を明示することが推奨されます。",
                category="実施可能要件",
                severity="中",
                location="発明の詳細な説明",
                suggestion="課題を解決するための手段に対する効果を整理して追記してください。",
            )
        )
    return findings


def check_support_requirement(doc: DocxDocument) -> List[Finding]:
    paragraphs = doc.paragraphs
    description_text = "\n".join(p for p in paragraphs if "請求項" not in p)
    claims = _extract_claims(paragraphs)
    findings: List[Finding] = []
    for number, idx, body in claims:
        keywords = _extract_keywords(body)
        missing = [kw for kw in keywords if kw and kw not in description_text]
        if not missing:
            continue
        snippet = _paragraph_snippet(paragraphs, idx)
        findings.append(
            Finding(
                identifier=f"SUP-{number:03d}",
                title=f"請求項{number}のサポート要件",
                message=f"請求項{number}の用語 {', '.join(missing[:5])} が明細書本文で十分に説明されていません。",
                category="サポート要件",
                severity="中",
                location=f"請求項{number}",
                context=snippet,
                suggestion="請求項の構成要素と対応する実施形態・作用効果を本文で詳述してください。",
            )
        )
    return findings


def check_claim_numbering(doc: DocxDocument) -> List[Finding]:
    paragraphs = doc.paragraphs
    claims = _extract_claims(paragraphs)
    findings: List[Finding] = []
    if not claims:
        return findings
    expected = 1
    for number, idx, _ in claims:
        if number != expected:
            snippet = _paragraph_snippet(paragraphs, idx)
            findings.append(
                Finding(
                    identifier=f"NUM-{number:03d}",
                    title="請求項番号の不整合",
                    message=f"請求項{number}が想定順序 {expected} とずれています。",
                    category="形式",
                    severity="中",
                    location=f"請求項{number}",
                    context=snippet,
                    suggestion="請求項番号が連続するように調整してください。",
                )
            )
            expected = number + 1
        else:
            expected += 1
    return findings


def check_claim_notation(doc: DocxDocument) -> List[Finding]:
    findings: List[Finding] = []
    paragraphs = doc.paragraphs
    pattern = re.compile(r"請求項([0-9]{1,2})")
    for idx, para in enumerate(paragraphs):
        for match in pattern.finditer(para):
            ascii_number = match.group(1)
            fullwidth_number = ascii_number.translate({
                ord("0"): "０",
                ord("1"): "１",
                ord("2"): "２",
                ord("3"): "３",
                ord("4"): "４",
                ord("5"): "５",
                ord("6"): "６",
                ord("7"): "７",
                ord("8"): "８",
                ord("9"): "９",
            })
            original = f"請求項{ascii_number}"
            replacement = f"請求項{fullwidth_number}"

            def _make_fix(old: str, new: str) -> AutoFix:
                return AutoFix(
                    label="請求項番号を全角に変換",
                    apply=lambda d, o=old, n=new: d.replace_text(o, n),
                )

            findings.append(
                Finding(
                    identifier=f"FMT-{idx:03d}-{ascii_number}",
                    title="請求項表記の統一",
                    message=f"『{original}』が半角数字になっています。和文特許では全角数字が一般的です。",
                    category="形式",
                    severity="低",
                    location=f"段落{idx + 1}",
                    context=_paragraph_snippet(paragraphs, idx),
                    suggestion="請求項番号を全角に統一してください。",
                    autofix=_make_fix(original, replacement),
                )
            )
    parentheses_pattern = re.compile(r"([一-龥ぁ-んァ-ン])\(([^)]+)\)")
    for idx, para in enumerate(paragraphs):
        for match in parentheses_pattern.finditer(para):
            original = match.group(0)
            replacement = original.replace("(", "（").replace(")", "）")

            def _make_fix(old: str, new: str) -> AutoFix:
                return AutoFix(
                    label="全角括弧へ置換",
                    apply=lambda d, o=old, n=new: d.replace_text(o, n),
                )

            findings.append(
                Finding(
                    identifier=f"FMT-PAREN-{idx:03d}-{match.start()}",
                    title="括弧の全角化",
                    message="和文中で半角括弧が使われています。公用文では全角括弧が推奨されます。",
                    category="形式",
                    severity="低",
                    location=f"段落{idx + 1}",
                    context=_paragraph_snippet(paragraphs, idx),
                    suggestion="全角括弧（ ）を使用してください。",
                    autofix=_make_fix(original, replacement),
                )
            )
    return findings


def check_placeholder_text(doc: DocxDocument) -> List[Finding]:
    paragraphs = doc.paragraphs
    patterns = [
        (r"TODO", "ドラフトのプレースホルダー 'TODO' が残っています。"),
        (r"XXXX", "仮文字 'XXXX' が残っています。"),
        (r"\?\?\?", "疑問符が連続しており、未決箇所の可能性があります。"),
        (r"要修正", "『要修正』というメモが残っています。"),
    ]
    findings: List[Finding] = []
    for idx, para in enumerate(paragraphs):
        for pattern, message in patterns:
            if re.search(pattern, para, re.IGNORECASE):
                findings.append(
                    Finding(
                        identifier=f"TMP-{idx:03d}",
                        title="プレースホルダーの除去",
                        message=message,
                        category="品質",
                        severity="高",
                        location=f"段落{idx + 1}",
                        context=_paragraph_snippet(paragraphs, idx),
                        suggestion="該当箇所を最終文言に置き換えてください。",
                    )
                )
    return findings


def check_figure_consistency(doc: DocxDocument) -> List[Finding]:
    paragraphs = doc.paragraphs
    figures_in_text = set(re.findall(r"図[0-9０-９]+", doc.full_text))
    findings: List[Finding] = []
    if not figures_in_text:
        return findings
    description_section = any("図面の簡単な説明" in p for p in paragraphs)
    if not description_section:
        findings.append(
            Finding(
                identifier="FIG-001",
                title="図面説明の不足",
                message="図の参照があるにもかかわらず『図面の簡単な説明』が確認できません。",
                category="図面",
                severity="中",
                location="図面関連",
                suggestion="各図について図面の簡単な説明を追記してください。",
            )
        )
    return findings


# === チェックの実行 ===========================================================


CheckFunction = Callable[[DocxDocument], List[Finding]]


CHECKS: Sequence[CheckFunction] = (
    check_required_sections,
    check_enablement,
    check_support_requirement,
    check_claim_numbering,
    check_claim_notation,
    check_placeholder_text,
    check_figure_consistency,
)


def run_checks(doc: DocxDocument) -> List[Finding]:
    findings: List[Finding] = []
    for check in CHECKS:
        try:
            findings.extend(check(doc))
        except Exception as exc:  # pragma: no cover - チェックの想定外エラーを捕捉
            findings.append(
                Finding(
                    identifier=f"ERR-{check.__name__}",
                    title="チェックエラー",
                    message=f"{check.__name__} の実行中にエラーが発生しました: {exc}",
                    category="システム",
                    severity="高",
                    location="-",
                    suggestion="スクリプトの開発者へ連絡してください。",
                )
            )
    # identifier + message で重複排除
    unique: Dict[Tuple[str, str], Finding] = {}
    for finding in findings:
        unique[(finding.identifier, finding.message)] = finding
    return list(unique.values())


# === インタラクティブ処理 =====================================================


def _prompt(prompt: str) -> str:
    try:
        return input(prompt)
    except EOFError:  # 非対話環境での安全策
        return ""


def review_findings(findings: List[Finding], doc: DocxDocument) -> List[FindingResolution]:
    resolutions: List[FindingResolution] = []
    if not findings:
        print("指摘事項はありませんでした。")
        return resolutions
    print(f"\n検出した指摘: {len(findings)} 件")
    for idx, finding in enumerate(findings, 1):
        print("\n" + "=" * 80)
        print(f"[{idx}] {finding.title} ({finding.category} / 重要度:{finding.severity})")
        print(textwrap.fill(finding.message, width=78))
        if finding.context:
            print(f"  └ 文脈: {finding.context}")
        if finding.suggestion:
            print("提案:")
            print(textwrap.indent(textwrap.fill(finding.suggestion, width=74), "  - "))

        applied_fix = False
        status = "未対応"

        if finding.autofix is not None:
            answer = _prompt("自動修正を適用しますか？ [y]はい / [n]いいえ / [s]後で : ").strip().lower()
            if answer == "y":
                applied_fix = finding.autofix.apply(doc)
                status = "自動修正済"
            elif answer == "s":
                status = "要確認"
            else:
                status = "手動対応予定"
        else:
            answer = _prompt("対応状況を入力してください。 [d]対応予定 / [s]後で / [i]無視 : ").strip().lower()
            if answer == "d":
                status = "手動対応予定"
            elif answer == "i":
                status = "一時保留"
            else:
                status = "要確認"

        note = _prompt("備考（Enterでスキップ）: ").strip()
        resolutions.append(
            FindingResolution(
                finding=finding,
                status=status,
                note=note,
                applied_fix=applied_fix,
            )
        )
    return resolutions


def generate_report(resolutions: List[FindingResolution], output_path: str, original_path: str) -> None:
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    lines: List[str] = []
    timestamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    lines.append(f"# 明細書チェックレポート\n")
    lines.append(f"- 生成日時: {timestamp}")
    lines.append(f"- 元ファイル: {os.path.basename(original_path)}")
    lines.append(f"- 指摘数: {len(resolutions)}\n")
    for res in resolutions:
        finding = res.finding
        lines.append(f"## {finding.title} ({finding.category})")
        lines.append(f"- 重要度: {finding.severity}")
        lines.append(f"- 識別子: {finding.identifier}")
        lines.append(f"- 場所: {finding.location}")
        lines.append(f"- 状況: {res.status}")
        if res.applied_fix:
            lines.append("- 対応: 自動修正を適用")
        if finding.context:
            lines.append(f"- 文脈: {finding.context}")
        lines.append("")
        lines.append(textwrap.indent(textwrap.fill(finding.message, width=78), "> "))
        if finding.suggestion:
            lines.append("")
            lines.append("**推奨対応**")
            lines.append(textwrap.fill(finding.suggestion, width=78))
        if res.note:
            lines.append("")
            lines.append("**備考**")
            lines.append(textwrap.fill(res.note, width=78))
        lines.append("")
    with open(output_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))
    print(f"レポートを出力しました: {output_path}")


# === CLI ======================================================================


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="特許明細書案を解析して指摘一覧を出力するツール")
    parser.add_argument("input", nargs="?", help="解析する .docx ファイルのパス")
    parser.add_argument("--output", help="修正後のdocxファイル出力先")
    parser.add_argument("--report", help="レポート（Markdown）の出力先")
    parser.add_argument(
        "--non-interactive",
        action="store_true",
        help="自動修正を適用せずレポートのみ生成する",
    )
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    input_path = args.input
    if not input_path:
        input_path = _prompt("解析する .docx ファイルのパスを入力してください: ").strip()
    if not input_path:
        print("入力ファイルが指定されませんでした。処理を終了します。")
        return 1
    if not os.path.exists(input_path):
        print(f"ファイルが見つかりません: {input_path}")
        return 1
    try:
        doc = DocxDocument(input_path)
    except Exception as exc:
        print(f"docx の読み込みに失敗しました: {exc}")
        return 1

    findings = run_checks(doc)
    if args.non_interactive:
        resolutions = [
            FindingResolution(finding=f, status="要確認", note="", applied_fix=False)
            for f in findings
        ]
    else:
        resolutions = review_findings(findings, doc)

    any_fix = any(res.applied_fix for res in resolutions)
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    output_docx = args.output or os.path.join(
        os.path.dirname(input_path), f"{base_name}_revised.docx"
    )
    report_path = args.report or os.path.join(
        os.path.dirname(input_path), f"{base_name}_check_report.md"
    )

    if any_fix:
        doc.save(output_docx)
        print(f"自動修正を適用したファイルを保存しました: {output_docx}")
    else:
        print("自動修正は適用されませんでした。原本は変更されていません。")

    generate_report(resolutions, report_path, input_path)
    return 0


if __name__ == "__main__":  # pragma: no cover - CLI エントリポイント
    sys.exit(main())


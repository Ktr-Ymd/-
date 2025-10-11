from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from io import BytesIO
from typing import List

from flask import Flask, redirect, render_template, request, url_for
from PyPDF2 import PdfReader
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


@dataclass
class PatentDocument:
    name: str
    page_count: int
    chunks: List[str]
    vectorizer: TfidfVectorizer
    tfidf_matrix: any = field(repr=False)

    def search(self, query: str, limit: int = 5) -> List[tuple[int, float]]:
        if not query.strip():
            return []
        query_vec = self.vectorizer.transform([query])
        similarities = cosine_similarity(query_vec, self.tfidf_matrix)[0]
        scored_indices = sorted(
            enumerate(similarities), key=lambda item: item[1], reverse=True
        )
        return scored_indices[:limit]


def create_app() -> Flask:
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB uploads

    store: dict[str, PatentDocument | None] = {"current": None}

    def _chunk_text(text: str, chunk_size: int = 800, overlap: int = 200) -> List[str]:
        cleaned = re.sub(r"\s+", " ", text).strip()
        if not cleaned:
            return []
        chunks: List[str] = []
        start = 0
        while start < len(cleaned):
            end = min(len(cleaned), start + chunk_size)
            chunk = cleaned[start:end]
            chunks.append(chunk)
            if end == len(cleaned):
                break
            start = max(end - overlap, start + 1)
        return chunks

    def _build_document(filename: str, file_bytes: bytes) -> PatentDocument:
        reader = PdfReader(BytesIO(file_bytes))
        text_parts: List[str] = []
        for page in reader.pages:
            page_text = page.extract_text() or ""
            text_parts.append(page_text)
        text = "\n".join(text_parts)
        chunks = _chunk_text(text)
        if not chunks:
            chunks = ["(PDFからテキストを抽出できませんでした)"]
        vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(3, 5))
        tfidf_matrix = vectorizer.fit_transform(chunks)
        return PatentDocument(
            name=filename,
            page_count=len(reader.pages),
            chunks=chunks,
            vectorizer=vectorizer,
            tfidf_matrix=tfidf_matrix,
        )

    @app.route("/", methods=["GET", "POST"])
    def index():
        if request.method == "POST":
            file = request.files.get("pdf")
            if not file or file.filename == "":
                return redirect(url_for("index"))
            filename = file.filename
            data = file.read()
            try:
                document = _build_document(filename, data)
            except Exception:
                document = None
            store["current"] = document
            return redirect(url_for("index"))

        document = store.get("current")
        query = request.args.get("query", "")
        results: List[tuple[int, float]] = []
        if document and query:
            results = document.search(query)
        return render_template(
            "index.html",
            document=document,
            results=results,
            query=query,
            math=math,
        )

    return app


app = create_app()

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)

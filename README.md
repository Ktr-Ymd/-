# 特許公報リーダー

PDF形式の特許公報を読み込み、文書内の記述のみを検索・閲覧できる簡易Webアプリです。アップロードした公報以外の情報は表示されません。

## セットアップ

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 実行方法

```bash
flask --app app run --host 0.0.0.0 --port 5000
```

ブラウザで <http://localhost:5000> を開いてください。

1. PDFの特許公報をアップロードすると、全文を分割した結果と基本情報が表示されます。
2. 検索欄にキーワードを入力すると、ベクトル検索（RAG）で類似度の高いセクションが返されます。

## 技術的ポイント

- PyPDF2で公報PDFからテキストを抽出
- 文字n-gramのTF-IDFベクトルでセクションをベクトル化
- コサイン類似度による簡易RAG検索
- Flask + Jinja2でシンプルなUIを構築

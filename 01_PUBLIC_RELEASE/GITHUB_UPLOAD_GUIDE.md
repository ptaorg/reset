# GITHUB_UPLOAD_GUIDE

## GitHub Pagesへ載せる場合

1. `03_HTML_SITE/pta-starter-kit-v0.2.2-rich-site.zip` を展開する
2. 中身を `docs/starter-kit/` などに配置する
3. `index.html` が入口になるようにする
4. `.nojekyll` をリポジトリの公開ルートに置く
5. `robots.txt` と `sitemap.xml` のURLを実際の公開URLに合わせて修正する
6. ブラウザで `index.html`、`manual.html`、`templates.html`、`sources.html` を確認する

## 注意

GitHub Pagesでは日本語ファイル名より英数字ファイル名が安定します。HTML版は英数字ファイル名で作成済みです。

## 推奨配置例

```text
/
├── index.html
├── docs/
│   └── starter-kit/
│       ├── index.html
│       ├── manual.html
│       ├── templates.html
│       ├── checklist.html
│       ├── sources.html
│       └── template-pages/
└── .nojekyll
```

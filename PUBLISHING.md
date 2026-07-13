# 公開手順書

このドキュメントは、Todoリマインダーアプリを公開するための手順をまとめたものです。

## ✅ 完了した準備項目

- [x] README.md - プロジェクト概要と使用方法
- [x] LICENSE - MIT ライセンス
- [x] PRIVACY.md - プライバシーポリシー
- [x] TERMS.md - 利用規約
- [x] package.json - npm パッケージ情報
- [x] docs/spec.md - 仕様書
- [x] .gitignore - 公開不要ファイルの除外

## 🚀 GitHub でのリポジトリ公開

### 1. リポジトリをパブリックに設定

1. GitHub にログイン
2. リポジトリページへアクセス: https://github.com/Coffeeplantation/todo-reminder-app
3. **Settings** → **Repository** を開く
4. **Visibility** セクションで **Change visibility** をクリック
5. **Public** を選択し、確認

### 2. GitHub Pages で無料ホスティング

1. リポジトリの **Settings** → **Pages** を開く
2. **Source** を **main/master branch** に設定
3. **Save** をクリック
4. URL が生成されます（`https://Coffeeplantation.github.io/todo-reminder-app`）

## 📦 Chrome Webストアでの公開（オプション）

### 必要な準備物
1. Google Developer Account（$5の登録料）
2. アプリのアイコン（128x128 pixel以上）
3. スクリーンショット（5枚以上）
4. 短い説明文（最大132文字）
5. 詳細な説明（最大4000文字）
6. プライバシーポリシーの URL

### 公開手順
1. [Chrome Webストア デベロッパーダッシュボード](https://chrome.google.com/webstore/devconsole) へアクセス
2. **新しいアイテム** をクリック
3. マニフェストファイル（manifest.json）をアップロード
4. ストア情報を入力
5. 審査に提出

## 📝 必要なマニフェストファイル（manifest.json）

Chrome Webストア用に以下のファイルを作成：

```json
{
  "manifest_version": 3,
  "name": "Todoリマインダー",
  "version": "1.0.0",
  "description": "タスクボードと予定管理を統合したシンプルで強力なタスク管理アプリ",
  "permissions": ["storage", "scripting"],
  "action": {
    "default_popup": "browser/index.html",
    "default_title": "Todoリマインダー"
  },
  "icons": {
    "128": "icons/icon-128.png"
  }
}
```

## 🖼️ 公開用素材

以下のファイルを用意：
- **icons/icon-128.png** - Chrome Webストア用アイコン
- **icons/icon-64.png** - 補助アイコン
- **スクリーンショット** - 主な機能を示す画像5枚

既存のスクリーンショット：
- board-header2.png
- card-open.png
- card-overflow.png
- current-schedule-modal.png
- settings-ollama.png

## 📚 ドキュメント確認チェックリスト

- [x] README.md が詳細で分かりやすい
- [x] インストール手順が明確
- [x] LICENSE ファイルが含まれている
- [x] プライバシーポリシーがある
- [x] 利用規約がある
- [x] コントリビューション情報がある
- [x] 問題報告方法が記載されている

## 🔍 検証

### リポジトリの確認
```bash
# GitHub リポジトリを確認
git remote -v

# ローカル設定が含まれていないか確認
cat .gitignore
```

### アプリの動作確認
```bash
# ブラウザ版を実行
cd browser
python -m http.server 8000
# http://localhost:8000 でアクセス
```

## 📊 公開後の推奨事項

1. **README 更新**: 実際に公開されたリンクを記載
2. **ソーシャルメディア**: GitHub ページへのリンクを共有
3. **Issue テンプレート**: バグ報告用テンプレートを作成
4. **Release ノート**: GitHub Releases で主な機能をまとめる
5. **定期的な更新**: 機能追加や改善を継続

## 🆘 トラブルシューティング

### GitHub Pages が表示されない場合
- Settings → Pages でブランチが正しく設定されているか確認
- index.html が公開されているか確認

### Chrome Webストア審査に落ちた場合
- プライバシーポリシーが完備されているか確認
- マニフェストファイルが正しいか確認
- 規約を遵守しているか確認

---

**更新日**: 2026-07-14

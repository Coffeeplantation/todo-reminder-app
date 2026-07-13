# GitHub リポジトリ公開設定スクリプト
# 実行前に以下を確認してください：
# 1. GitHub CLI がインストールされていること（https://cli.github.com/）
# 2. gh auth login で GitHub に認証されていること

Write-Host "🚀 GitHub リポジトリ公開設定を開始します..." -ForegroundColor Cyan

# GitHub CLI の確認
$ghPath = Where-Object { $_.Name -eq "gh.exe" } -InputObject @(Get-Command gh -ErrorAction SilentlyContinue)
if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    Write-Host "❌ GitHub CLI がインストールされていません" -ForegroundColor Red
    Write-Host ""
    Write-Host "GitHub CLI をインストールしてください："
    Write-Host "  1. https://cli.github.com/ から CLI をダウンロード"
    Write-Host "  2. インストール後、PowerShell を再起動"
    Write-Host "  3. gh auth login で GitHub に認証"
    Write-Host "  4. このスクリプトを再実行"
    exit 1
}

Write-Host "✅ GitHub CLI が見つかりました" -ForegroundColor Green

# 認証確認
$auth = gh auth status 2>&1
if ($auth -match "not authenticated") {
    Write-Host "❌ GitHub に認証されていません" -ForegroundColor Red
    Write-Host ""
    Write-Host "認証してください："
    Write-Host "  gh auth login"
    exit 1
}

Write-Host "✅ GitHub に認証されています" -ForegroundColor Green
Write-Host ""

# リポジトリ情報
$owner = "Coffeeplantation"
$repo = "todo-reminder-app"
$repoUrl = "$owner/$repo"

Write-Host "リポジトリ: $repoUrl" -ForegroundColor Cyan

# 1. リポジトリを公開（private → public）
Write-Host ""
Write-Host "1️⃣  リポジトリをパブリックに設定中..." -ForegroundColor Yellow
gh repo edit $repoUrl --visibility public
Write-Host "✅ リポジトリをパブリックに設定しました" -ForegroundColor Green

# 2. Description を設定
Write-Host ""
Write-Host "2️⃣  Description を設定中..." -ForegroundColor Yellow
gh repo edit $repoUrl `
  --description "タスクボードと予定管理を統合したシンプルで強力なタスク管理アプリ。ブラウザで動作し、オフラインでも使用可能。" `
  --homepage "https://github.com/Coffeeplantation/todo-reminder-app"
Write-Host "✅ Description を設定しました" -ForegroundColor Green

# 3. Topics を設定
Write-Host ""
Write-Host "3️⃣  Topics を設定中..." -ForegroundColor Yellow
gh repo edit $repoUrl `
  --add-topic "task-management" `
  --add-topic "kanban" `
  --add-topic "scheduler" `
  --add-topic "todo-list" `
  --add-topic "japanese"
Write-Host "✅ Topics を設定しました" -ForegroundColor Green

# 4. Release を作成
Write-Host ""
Write-Host "4️⃣  v1.0.0 Release を作成中..." -ForegroundColor Yellow

$releaseBody = @"
# Todoリマインダー v1.0.0

## 🎉 初回リリース

タスクボードと予定管理を統合したシンプルで強力なタスク管理アプリです。

## 🌟 主な機能

- **複数ボード管理**: 月別・プロジェクト別など、複数のタスクボードを同時管理
- **予定インポート**: Excel、CSV、カレンダー（.ics）、テキスト、URL から予定を取り込み
- **AI解析**: Claude API または Ollama を使用した自動予定抽出
- **現行予定表**: ボード内のタスクとインポート予定を一元表示
- **ドラッグ&ドロップ**: タスク移動による状態管理
- **ローカルストレージ**: すべてのデータをブラウザに保存（サーバー不要）

## 🚀 使い方

### ブラウザ版
\`\`\`bash
cd browser
python -m http.server 8000
# http://localhost:8000 を開く
\`\`\`

### Excel 版（VBA）
\`Todoリマインダー.xlsm\` を開く

## 📋 システム要件
- モダンなウェブブラウザ（Chrome、Firefox、Safari等）
- Python 3.x（サーバー起動用）

## 🔧 セットアップ

1. リポジトリをクローン
   \`\`\`bash
   git clone https://github.com/Coffeeplantation/todo-reminder-app.git
   \`\`\`

2. ブラウザ版を起動
   \`\`\`bash
   cd browser
   python -m http.server 8000
   \`\`\`

3. ローカルAI（オプション）
   - [Ollama](https://ollama.ai) をインストール
   - \`ollama pull qwen2.5:3b\`

## 📝 ドキュメント
- [README](./README.md) - 使用方法
- [PRIVACY](./PRIVACY.md) - プライバシーポリシー
- [CONTRIBUTING](./CONTRIBUTING.md) - コントリビューション

## 📄 ライセンス
MIT License

---

**公開日**: 2026-07-14
"@

gh release create v1.0.0 `
  --title "Todoリマインダー v1.0.0" `
  --notes $releaseBody `
  --latest
Write-Host "✅ v1.0.0 Release を作成しました" -ForegroundColor Green

# 完了
Write-Host ""
Write-Host "🎉 すべての設定が完了しました！" -ForegroundColor Cyan
Write-Host ""
Write-Host "✅ リポジトリをパブリックに設定"
Write-Host "✅ Description と Topics を追加"
Write-Host "✅ v1.0.0 Release を作成"
Write-Host ""
Write-Host "リポジトリはこちらから確認できます："
Write-Host "  https://github.com/$repoUrl" -ForegroundColor Blue
Write-Host ""

# GitHub リポジトリ確認・設定スクリプト

Write-Host "🔍 GitHub リポジトリの状態を確認中..." -ForegroundColor Cyan
Write-Host ""

# GitHub CLI が利用可能か確認
if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    Write-Host "⚠️  GitHub CLI がインストールされていません" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "以下の手順に従って GitHub CLI をインストールしてください：" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "1. PowerShell を管理者として実行"
    Write-Host "2. 以下のコマンドを実行："
    Write-Host "   choco install gh"
    Write-Host "   # または"
    Write-Host "   winget install GitHub.CLI"
    Write-Host ""
    Write-Host "3. インストール後、PowerShell を再起動して gh auth login を実行"
    Write-Host ""
    exit 1
}

# リポジトリ情報
$owner = "Coffeeplantation"
$repo = "todo-reminder-app"
$repoUrl = "$owner/$repo"

# リポジトリ情報を表示
Write-Host "📦 リポジトリ情報取得中..." -ForegroundColor Cyan
$repoInfo = gh api repos/$repoUrl --jq '.private,.description,.topics,.homepage'

Write-Host "リポジトリ: $repoUrl"
Write-Host ""

# リポジトリを公開
Write-Host "🔓 リポジトリをパブリックに設定中..." -ForegroundColor Yellow
gh repo edit $repoUrl --visibility public
Write-Host "✅ リポジトリをパブリックに設定しました" -ForegroundColor Green
Write-Host ""

# Description を設定
Write-Host "📝 Description を設定中..." -ForegroundColor Yellow
gh repo edit $repoUrl `
  --description "タスクボードと予定管理を統合したタスク管理アプリ。ブラウザで動作し、オフラインでも使用可能。"
Write-Host "✅ Description を設定しました" -ForegroundColor Green
Write-Host ""

# Topics を設定
Write-Host "🏷️  Topics を設定中..." -ForegroundColor Yellow
gh repo edit $repoUrl `
  --add-topic "task-management" `
  --add-topic "kanban" `
  --add-topic "scheduler" `
  --add-topic "todo-list" `
  --add-topic "japanese"
Write-Host "✅ Topics を設定しました" -ForegroundColor Green
Write-Host ""

# Release を確認・作成
Write-Host "🏷️  Release を確認中..." -ForegroundColor Yellow
$releases = gh release list -R $repoUrl --limit 1 2>$null
if ($releases -match "v1.0.0") {
    Write-Host "ℹ️  v1.0.0 Release は既に存在します" -ForegroundColor Blue
} else {
    Write-Host "📝 v1.0.0 Release を作成中..." -ForegroundColor Yellow

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

## 🚀 クイックスタート

### ブラウザ版を実行
\`\`\`bash
cd browser
python -m http.server 8000
# http://localhost:8000 を開く
\`\`\`

### ローカルAI（オプション）
- [Ollama](https://ollama.ai) をインストール
- \`ollama pull qwen2.5:3b\`

## 📋 ドキュメント
- [README](./README.md)
- [PRIVACY](./PRIVACY.md)
- [CONTRIBUTING](./CONTRIBUTING.md)

## 📄 ライセンス
MIT License

---
**公開日**: 2026-07-14
"@

    gh release create v1.0.0 `
      --title "Todoリマインダー v1.0.0" `
      --notes $releaseBody `
      --latest `
      -R $repoUrl

    Write-Host "✅ v1.0.0 Release を作成しました" -ForegroundColor Green
}

Write-Host ""
Write-Host "🎉 すべての設定が完了しました！" -ForegroundColor Cyan
Write-Host ""
Write-Host "✅ リポジトリをパブリックに設定"
Write-Host "✅ Description を追加"
Write-Host "✅ Topics を追加"
Write-Host "✅ Release を作成"
Write-Host ""
Write-Host "以下のリンクで確認できます：" -ForegroundColor Cyan
Write-Host "  リポジトリ: https://github.com/$repoUrl" -ForegroundColor Blue
Write-Host "  Releases: https://github.com/$repoUrl/releases" -ForegroundColor Blue
Write-Host ""
Write-Host "GitHub Pages が有効化されるまで約10分待機してください。" -ForegroundColor Yellow
Write-Host "  Pages: https://$owner.github.io/$repo" -ForegroundColor Blue
Write-Host ""

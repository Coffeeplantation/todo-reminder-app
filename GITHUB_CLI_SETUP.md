# GitHub CLI セットアップガイド

リポジトリを自動的に公開設定するために、GitHub CLI をインストールして実行します。

## 📦 GitHub CLI インストール

### Windows での インストール

#### 方法 1: Chocolatey（推奨）
```powershell
# PowerShell を管理者として実行
choco install gh
```

#### 方法 2: WinGet
```powershell
# PowerShell を管理者として実行
winget install GitHub.CLI
```

#### 方法 3: 直接ダウンロード
1. https://github.com/cli/cli/releases から最新版をダウンロード
2. インストーラーを実行

### macOS でのインストール
```bash
brew install gh
```

### Linux でのインストール
```bash
# Ubuntu/Debian
sudo apt install gh

# Fedora/RHEL
sudo dnf install gh

# Arch Linux
sudo pacman -S github-cli
```

## 🔐 GitHub に認証

### 1. 認証コマンドを実行
```powershell
gh auth login
```

### 2. プロンプトに答える
```
? What is your preferred protocol for Git operations?  [Use arrows to move, type to filter]
> HTTPS
  SSH

? Authenticate Git with your GitHub credentials? (Y/n) 
> Y

? How would you like to authenticate GitHub CLI?  [Use arrows to move, type to filter]
> Login with a web browser
  Paste an authentication token
```

### 3. ブラウザで認証
- ブラウザが自動で開きます
- "Authorize github-cli" をクリック
- 認証コードを入力（要求されれば）

### 4. 完了
```
✓ Authentication complete. You're logged in as [username].
```

## ✅ 認証確認

```powershell
gh auth status
```

以下のように表示されたら成功：
```
github.com
  ✓ Logged in to github.com as Coffeeplantation
  ✓ Git operations for github.com configured to use https protocol.
  ✓ Token: gho_****...
```

## 🚀 公開設定スクリプトを実行

### 1. PowerShell を起動（管理者権限不要）
```powershell
cd C:\Users\katsu\Documents\todo-reminder-app
```

### 2. スクリプト実行権限を設定
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

確認が表示されたら `Y` を入力

### 3. スクリプトを実行
```powershell
.\setup-github-public.ps1
```

### 実行結果

スクリプトが以下を自動実行します：

```
🚀 GitHub リポジトリ公開設定を開始します...
✅ GitHub CLI が見つかりました
✅ GitHub に認証されています

リポジトリ: Coffeeplantation/todo-reminder-app

1️⃣  リポジトリをパブリックに設定中...
✅ リポジトリをパブリックに設定しました

2️⃣  Description を設定中...
✅ Description を設定しました

3️⃣  Topics を設定中...
✅ Topics を設定しました

4️⃣  v1.0.0 Release を作成中...
✅ v1.0.0 Release を作成しました

🎉 すべての設定が完了しました！
```

## 🔍 確認

完了後、以下で確認できます：

1. **リポジトリ**: https://github.com/Coffeeplantation/todo-reminder-app
   - Visibility が "Public" になっている
   - Description が表示されている
   - Topics が表示されている

2. **Releases**: https://github.com/Coffeeplantation/todo-reminder-app/releases
   - v1.0.0 が作成されている
   - リリースノートが表示されている

3. **GitHub Pages**: https://Coffeeplantation.github.io/todo-reminder-app
   - アプリが表示されている（設定後10分程度待機）

## 🆘 トラブルシューティング

### "GitHub CLI not found"
```powershell
# インストールを確認
gh version
```

見つからない場合は上記のインストール手順を実行

### "not authenticated"
```powershell
gh auth login
```

再度認証を実行

### "permission denied"
```powershell
# PowerShell 実行権限を確認
Get-ExecutionPolicy

# RemoteSigned に設定
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### スクリプト実行エラー
```powershell
# スクリプトをテスト実行
.\setup-github-public.ps1 -WhatIf
```

## 📚 参考リンク

- [GitHub CLI 公式ドキュメント](https://cli.github.com/manual/)
- [gh repo edit コマンド](https://cli.github.com/manual/gh_repo_edit)
- [gh release create コマンド](https://cli.github.com/manual/gh_release_create)

---

**更新日**: 2026-07-14

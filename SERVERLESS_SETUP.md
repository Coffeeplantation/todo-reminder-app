# 🚀 サーバーレス基盤セットアップガイド

サーバーレス環境でのメール自動化システムを構築します。

## 📦 必要なサービス（すべて無料枠あり）

- **Supabase** - PostgreSQL ホスティング
- **Vercel** - サーバーレス関数
- **SendGrid** - メール送信
- **GitHub Actions** - 定期実行

---

## 1️⃣ **Supabase セットアップ**

### ステップ1：Supabase プロジェクト作成

1. https://supabase.com にアクセス
2. **"Sign up"** をクリック
3. GitHub アカウントで登録（Coffeeplantation）
4. **"New project"** をクリック
5. プロジェクト名：`todo-reminder-prod`
6. 地域：`Southeast Asia (Singapore)` または日本に近い地域
7. パスワードを設定（メモしておく）
8. **"Create new project"** をクリック

### ステップ2：テーブル作成

1. Supabase ダッシュボードで **"SQL Editor"** をクリック
2. 新しいクエリを開いて、以下を実行：

```sql
-- タスクテーブル
CREATE TABLE public.tasks (
  id BIGSERIAL PRIMARY KEY,
  name TEXT NOT NULL,
  due_date DATE NOT NULL,
  email TEXT NOT NULL,
  participants TEXT[] DEFAULT '{}',
  status TEXT DEFAULT 'pending',
  board_id TEXT,
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW()
);

-- 送信履歴テーブル
CREATE TABLE public.email_history (
  id BIGSERIAL PRIMARY KEY,
  task_id BIGINT REFERENCES public.tasks(id),
  email TEXT NOT NULL,
  subject TEXT,
  body TEXT,
  status TEXT DEFAULT 'sent',
  error_message TEXT,
  sent_at TIMESTAMP DEFAULT NOW()
);

-- ユーザー設定テーブル
CREATE TABLE public.user_settings (
  id BIGSERIAL PRIMARY KEY,
  user_email TEXT PRIMARY KEY,
  reminder_days_before INTEGER DEFAULT 3,
  email_time TIME DEFAULT '09:00',
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW()
);

-- インデックス作成
CREATE INDEX idx_tasks_due_date ON public.tasks(due_date);
CREATE INDEX idx_tasks_email ON public.tasks(email);
CREATE INDEX idx_email_history_task_id ON public.email_history(task_id);
```

### ステップ3：API キーを取得

1. **"Settings"** → **"API"** をクリック
2. 以下のキーをコピー（後で使用）：
   - **Project URL**: `https://xxxxx.supabase.co`
   - **anon key**: `eyJhbGc...`
   - **service_role key**: `eyJhbGc...`

> ⚠️ **重要**: service_role key は秘密にしてください！

---

## 2️⃣ **Vercel セットアップ**

### ステップ1：Vercel プロジェクト作成

1. https://vercel.com にアクセス
2. GitHub アカウントでログイン
3. **"New Project"** をクリック
4. `todo-reminder-app` リポジトリを選択
5. **"Import"** をクリック

### ステップ2：環境変数設定

1. **"Settings"** → **"Environment Variables"** をクリック
2. 以下の環境変数を追加：

```
SUPABASE_URL=https://xxxxx.supabase.co
SUPABASE_SERVICE_KEY=eyJhbGc...（service_role key）
SENDGRID_API_KEY=SG.xxxxx...
ANTHROPIC_API_KEY=sk-ant-xxxxx...
```

### ステップ3：関数ディレクトリ作成

```bash
cd todo-reminder-app
mkdir -p api
```

---

## 3️⃣ **SendGrid セットアップ**

### ステップ1：SendGrid アカウント作成

1. https://sendgrid.com にアクセス
2. **"Sign up"** をクリック
3. メールアドレス（coffeeplantation100@gmail.com）で登録
4. メール確認

### ステップ2：API キー生成

1. **"Settings"** → **"API Keys"** をクリック
2. **"Create API Key"** をクリック
3. 名前：`todo-reminder-key`
4. 権限：**"Full Access"**
5. **"Create & Use"** をクリック
6. キーをコピー（`SG.xxxxx...`）

### ステップ3：送信者設定

1. **"Settings"** → **"Sender Authentication"** をクリック
2. メールアドレスを認証
3. 確認メールの指示に従う

> 無料枠：100通/日、3000通/月

---

## 4️⃣ **GitHub Actions セットアップ**

### ステップ1：ワークフローファイル作成

`.github/workflows/daily-reminder.yml` を作成：

```yaml
name: Daily Email Reminder

on:
  schedule:
    - cron: '0 9 * * *'  # 毎日午前9時 UTC
  workflow_dispatch:     # 手動実行用

jobs:
  send-reminders:
    runs-on: ubuntu-latest
    steps:
      - name: Trigger Vercel Function
        env:
          VERCEL_TOKEN: ${{ secrets.VERCEL_TOKEN }}
        run: |
          curl -X POST \
            https://todo-reminder-app.vercel.app/api/send-reminders \
            -H "Authorization: Bearer $VERCEL_TOKEN" \
            -H "Content-Type: application/json" \
            -d '{}'
```

### ステップ2：GitHub Secrets 設定

1. リポジトリ → **"Settings"** → **"Secrets and variables"** → **"Actions"**
2. **"New repository secret"** をクリック
3. 名前：`VERCEL_TOKEN`
4. 値：Vercel プロジェクトから取得（Settings → Tokens）

---

## 📋 **チェックリスト**

### Supabase
- [ ] プロジェクト作成
- [ ] テーブル作成
- [ ] API キー取得
  - [ ] Project URL
  - [ ] anon key
  - [ ] service_role key

### Vercel
- [ ] プロジェクトインポート
- [ ] 環境変数設定
  - [ ] SUPABASE_URL
  - [ ] SUPABASE_SERVICE_KEY
  - [ ] SENDGRID_API_KEY
  - [ ] ANTHROPIC_API_KEY

### SendGrid
- [ ] アカウント作成
- [ ] API キー生成
- [ ] 送信者メール認証

### GitHub
- [ ] ワークフローファイル作成
- [ ] Vercel Token 設定

---

## ✅ **テスト方法**

すべてセットアップ後、以下でテスト：

```bash
# GitHub Actions を手動実行
# リポジトリ → Actions → "Daily Email Reminder" → "Run workflow"
```

または

```bash
# ローカルでテスト
curl -X POST http://localhost:3000/api/send-reminders \
  -H "Content-Type: application/json"
```

---

## 🆘 **トラブルシューティング**

### "Supabase connection failed"
- Project URL と API キーが正しいか確認
- Supabase ダッシュボードでテーブルが作成されているか確認

### "SendGrid auth failed"
- API キーが正しいか確認
- メール送信者が認証されているか確認

### "GitHub Actions not running"
- cron スケジュール形式が正しいか確認
- Vercel Token が設定されているか確認

---

## 📞 **次のステップ**

1. このガイドに従ってセットアップ
2. 各サービスのキーを記録
3. Vercel 関数コードを作成（次のドキュメント）
4. GitHub Actions で自動実行設定

---

**更新日**: 2026-07-14

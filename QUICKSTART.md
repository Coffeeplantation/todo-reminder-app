# ⚡ クイックスタート - サーバーレスメール自動化

最速でサーバーレス基盤をセットアップして、メール自動化をテストします。

## 🚀 3ステップで開始

### ステップ 1：ローカル開発環境構築（1分）

```bash
npm install
npm run setup-local
```

**実行内容：**
- ✅ SQLite DB を作成
- ✅ テーブルを自動生成
- ✅ テストデータを挿入
- ✅ .env.local ファイルを作成

### ステップ 2：API キー設定（3分）

```bash
# .env.local を開いて API キーを入力
# 対象: ANTHROPIC_API_KEY
```

取得方法：
- **Claude API**: https://console.anthropic.com/
  - 無料クレジット $5 付与（テスト用十分）
  - API キーをコピー

### ステップ 3：ローカルテスト実行（1分）

```bash
npm run test-local
```

**出力例：**
```
✅ DB に接続しました

1️⃣ タスク取得中...
   ✅ 3 件のタスクを取得

2️⃣ メール下書き生成テスト

   📝 タスク: プロジェクト企画書提出
      期日: 2026-07-15
      メール: test@example.com
      🤖 Claude API でメール生成中...
      📧 生成されたメール本文:
         お疲れ様です。
         プロジェクト企画書の提出期日が明日に迫っています。
         ...
```

---

## 📊 生成されるテストデータ

```
ID | タスク名 | 期日 | メール
1  | プロジェクト企画書提出 | 明日 | test@example.com
2  | クライアント打ち合わせ | 3日後 | test@example.com
3  | Q3 予算申請 | 本日 | test@example.com
```

---

## 🎯 本番環境へのステップ

### 1. Supabase アカウント作成

```bash
# 以下を実施（SERVERLESS_SETUP.md 参照）
1. https://supabase.com で無料登録
2. プロジェクト作成
3. SQL エディターでテーブル作成
4. API キー取得
```

### 2. Vercel へのデプロイ

```bash
# Vercel で環境変数設定
SUPABASE_URL=...
SUPABASE_SERVICE_KEY=...
SENDGRID_API_KEY=...
ANTHROPIC_API_KEY=...
```

### 3. GitHub Actions で自動実行

```bash
# リポジトリ → Actions で "Daily Email Reminder" を確認
# 毎日午前9時 UTC に自動実行開始
```

---

## 💰 コスト確認

```
┌──────────────┬────────────┬────────────┐
│ サービス     │ 無料枠     │ 月額コスト  │
├──────────────┼────────────┼────────────┤
│ Supabase     │ 無制限     │ $0         │
│ Vercel       │ 100GB帯域幅 │ $0         │
│ SendGrid     │ 3000通/月  │ $0         │
│ GitHub       │ 2000分/月  │ $0         │
│ Claude API   │ なし       │ 日本語文案 │
│              │            │ ≈ $0.03/日 │
├──────────────┼────────────┼────────────┤
│ 合計         │            │ ≈ $1/月    │
└──────────────┴────────────┴────────────┘
```

---

## 📚 ドキュメント

| ドキュメント | 説明 |
|-------------|------|
| **SERVERLESS_SETUP.md** | 本番環境の詳細セットアップ |
| **api/send-reminders.js** | Vercel 関数のコード |
| **vercel.json** | Vercel 設定ファイル |
| **.github/workflows/daily-reminder.yml** | GitHub Actions ワークフロー |

---

## ✅ チェックリスト

### ローカルテスト
- [ ] `npm run setup-local` を実行
- [ ] `.env.local` に Claude API キーを入力
- [ ] `npm run test-local` でメール生成をテスト

### 本番環境セットアップ
- [ ] Supabase プロジェクト作成
- [ ] Supabase テーブル作成
- [ ] SendGrid アカウント作成
- [ ] Vercel 環境変数設定
- [ ] GitHub Secrets 設定

### 本番環境テスト
- [ ] GitHub Actions で手動実行
- [ ] メール送信ログを確認
- [ ] Supabase で送信履歴を確認

---

## 🆘 トラブルシューティング

### "ANTHROPIC_API_KEY が未設定"
```bash
# .env.local に Claude API キーを追加
ANTHROPIC_API_KEY=sk-ant-xxxxx...
```

### "DB に接続できません"
```bash
# DB ファイルが破損している場合は削除して再構築
rm todo-reminder.db
npm run setup-local
```

### "メール生成に失敗"
- Claude API クレジットを確認
- API キーが正しいか確認
- インターネット接続を確認

---

## 🎓 学習ポイント

このセットアップで学べること：

1. **Serverless 関数**
   - Vercel Functions の使用方法
   - 環境変数の管理

2. **データベース**
   - SQLite でのローカル開発
   - Supabase での本番運用

3. **AI 統合**
   - Claude API の活用
   - ストリーミング応答の処理

4. **自動化**
   - GitHub Actions での定期実行
   - Cron ジョブの設定

5. **メール送信**
   - SendGrid の無料枠活用
   - テンプレート生成

---

**準備完了！** 🎉

さあ始めましょう：

```bash
npm install
npm run setup-local
npm run test-local
```

---

**更新日**: 2026-07-14

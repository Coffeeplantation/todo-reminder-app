# ログイン機能（リーダー／メンバー）と Supabase 設定

ブラウザ版（browser/index.html）に、起動時のログイン画面と、Supabase 経由の
タスクボード共有（リーダーが編集 → 全員が同じボードを見る）を追加しました。

**セットアップSQLは [supabase-setup.sql](./supabase-setup.sql) を SQL Editor に貼り付けて実行するだけです。**

## ロール（役割）

| ロール | 見える画面 |
| --- | --- |
| リーダー（`role = 'leader'`） | ヘッダー（アプリバー）を含む全機能。現行予定表・メニュー・設定など |
| メンバー（`role = 'member'`） | タスクボードのみ（ヘッダー非表示）。右下のチップからログアウト |

- ログイン状態はブラウザ（localStorage）に保存され、次回起動時も維持されます。
- リーダーは「☰ メニュー → 🚪 ログアウト」、メンバーは右下チップの「🚪 ログアウト」でログイン画面に戻れます。

## アカウントの管理場所

1. **Supabase 設定あり（オンライン時）** … Supabase の `app_users` 表で認証します。
2. **Supabase 未設定・オフライン時** … 内蔵アカウントにフォールバックします。
   - リーダー: `leader / 1234`
   - メンバー: `member / 1234`

※ パスワードは「とりあえず」の簡易運用です（平文保存）。本格運用時は Supabase Auth への移行を検討してください。

## Supabase のセットアップ手順

1. https://supabase.com でプロジェクトを作成する（無料枠でOK）。
2. SQL Editor で次を実行して、ユーザー表を作成する:

```sql
create table app_users (
  username     text primary key,
  password     text not null,          -- とりあえず平文（簡易運用）
  role         text not null default 'member',  -- 'leader' か 'member'
  display_name text
);

alter table app_users enable row level security;

-- anonキーでの読み取りを許可（ログイン照合に必要）
create policy "allow read for login" on app_users
  for select using (true);

-- 初期ユーザー
insert into app_users (username, password, role, display_name) values
  ('leader', '1234', 'leader', 'リーダー'),
  ('member', '1234', 'member', 'メンバー');
```

3. 「Project Settings → API」で **Project URL** と **anon public キー** をコピーする。
4. アプリにリーダーでログイン →「☰ メニュー → ⚙ 設定」で
   「🔐 ログイン：Supabase URL」「🔐 ログイン：Supabase anonキー」に貼り付ける。

または、git 管理外の `browser/config.local.js` に書いておくこともできます（設定が未入力のときだけ自動で取り込まれます）:

```js
window.LOCAL_CONFIG = {
  supabaseUrl: "https://xxxxx.supabase.co",
  supabaseAnonKey: "eyJhbGc...",
};
```

## ユーザーの追加・変更（アプリ内で完結）

リーダーでログイン →「☰ メニュー → 🔑 ログイン管理」から：

- **ユーザー一覧**: 表示名・役割（リーダー/メンバー）・パスワードをその場で編集して「保存」。
  「削除」でログイン不可に（自分自身は削除不可。履歴は残る）。
- **ユーザーを追加**: ユーザー名・表示名・役割・パスワードを入れて「追加」。
- **ログイン履歴（最新100件）**: 誰がいつログインしたか。「ログイン」（手動）と
  「自動」（保存済みセッションでのアプリ起動）を区別して記録。

Supabase の Table Editor（`app_users` / `login_history`）から直接編集することもできます。

## タスクボードの共有（同期）

Supabase の URL と anon キーを設定すると、ログイン認証に加えて **タスクボードの共有**も有効になります。

- **リーダー**の操作（タスク・ボード・名簿・配置などの変更）は約1.5秒後にまとめて
  `shared_state` 表（1行のJSON）へ保存されます。
- **全員**が約10秒ごと（＋ウィンドウを開いた時）に最新を読み込み、画面に反映します。
- 設定（APIキー・背景など）は共有されず、各自のブラウザに残ります。
- オフライン時・未設定時は従来どおり localStorage のみで動きます（壊れません）。

## 制限事項（今後の課題）

- メンバーのボード操作は保存されません（次の同期で リーダーのデータに戻ります）が、
  操作ボタン自体は現状押せます。完全な閲覧専用UIにする場合は追加実装が必要です。
- リーダーが複数人いて同時に編集した場合は「後から保存した方が勝ち」です。
- anon キーを知っていれば技術的には誰でも書き込めます（権限はアプリ側の制御）。
  厳密な権限管理が必要になったら Supabase Auth ＋ RLS への移行を検討してください。

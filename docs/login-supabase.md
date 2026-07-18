# ログイン機能（リーダー／メンバー）と Supabase 設定

ブラウザ版（browser/index.html）に、起動時のログイン画面を追加しました。

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

## ユーザーの追加・変更

Supabase の Table Editor で `app_users` に行を追加・編集するだけです。
`role` を `'leader'` にするとヘッダー付きの全機能、それ以外はメンバー（タスクボードのみ）になります。

## 制限事項（今後の課題）

- タスクデータ自体は各ブラウザの localStorage に保存されたままです（Supabase にはまだ同期しません）。
  別のPCのメンバーと同じボードを見るには、タスクデータの Supabase 同期が別途必要です。
- メンバーはヘッダー機能にアクセスできませんが、ボード上のタスク操作（移動・編集）は現状可能です。
  完全な閲覧専用にする場合は追加実装が必要です。

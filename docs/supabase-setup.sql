-- ============================================================
-- Todoリマインダー：Supabase セットアップSQL
-- Supabase ダッシュボード → SQL Editor に全文貼り付けて Run するだけ。
-- 何度実行しても安全（既にあれば作り直さない）。
-- ============================================================

-- 1) ログインアカウント表
--    role: 'leader'（ヘッダーあり・編集可） / 'member'（タスクボード閲覧のみ）
create table if not exists app_users (
  username     text primary key,
  password     text not null,                   -- とりあえず平文（簡易運用）
  role         text not null default 'member',  -- 'leader' か 'member'
  display_name text
);

alter table app_users enable row level security;

drop policy if exists "allow read for login" on app_users;
create policy "allow read for login" on app_users
  for select using (true);

insert into app_users (username, password, role, display_name) values
  ('leader', '1234', 'leader', 'リーダー'),
  ('member', '1234', 'member', 'メンバー')
on conflict (username) do nothing;

-- 2) タスクボード共有データ
--    アプリの状態（ボード・タスク・名簿など）を1行のJSONとして保持する。
--    リーダーの操作が保存され、メンバーは定期的に読み込んで表示する。
create table if not exists shared_state (
  id         text primary key,          -- 固定で 'main' の1行だけ使う
  data       jsonb not null,
  updated_at timestamptz not null default now()
);

alter table shared_state enable row level security;

drop policy if exists "read shared_state" on shared_state;
create policy "read shared_state" on shared_state
  for select using (true);

drop policy if exists "insert shared_state" on shared_state;
create policy "insert shared_state" on shared_state
  for insert with check (true);

drop policy if exists "update shared_state" on shared_state;
create policy "update shared_state" on shared_state
  for update using (true);

-- 3) ユーザー管理（リーダーがアプリ内からユーザーを追加・変更・削除できるようにする）
drop policy if exists "insert app_users" on app_users;
create policy "insert app_users" on app_users
  for insert with check (true);

drop policy if exists "update app_users" on app_users;
create policy "update app_users" on app_users
  for update using (true);

drop policy if exists "delete app_users" on app_users;
create policy "delete app_users" on app_users
  for delete using (true);

-- 4) ログイン履歴（リーダーが「誰がいつログインしたか」を確認できる）
create table if not exists login_history (
  id           bigint generated always as identity primary key,
  username     text not null,
  display_name text,
  role         text,
  kind         text,                              -- 'ログイン'（手動） / '自動'（保存済みセッションでの起動）
  logged_in_at timestamptz not null default now()
);

alter table login_history enable row level security;

drop policy if exists "read login_history" on login_history;
create policy "read login_history" on login_history
  for select using (true);

drop policy if exists "insert login_history" on login_history;
create policy "insert login_history" on login_history
  for insert with check (true);

-- 注意: anonキーを持つ人（＝アプリ利用者全員）は技術的には書き込みも可能です。
-- 「メンバーは閲覧のみ」「ユーザー管理はリーダーのみ」はアプリ側の制御です。
-- 簡易運用のための割り切りで、厳密な権限管理が必要になったら
-- Supabase Auth ＋ RLS への移行を検討してください。

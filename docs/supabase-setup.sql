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

-- 注意: anonキーを持つ人（＝アプリ利用者全員）は技術的には書き込みも可能です。
-- 「メンバーは閲覧のみ」はアプリ側の制御です。簡易運用のための割り切りで、
-- 厳密な権限管理が必要になったら Supabase Auth ＋ RLS への移行を検討してください。

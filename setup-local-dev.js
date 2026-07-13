#!/usr/bin/env node

/**
 * ローカル開発環境セットアップスクリプト
 *
 * 機能：
 * - SQLite DB 作成
 * - テスト用テーブル作成
 * - テストデータ挿入
 * - Vercel 関数のローカル実行環境構築
 */

const fs = require('fs');
const path = require('path');
const sqlite3 = require('sqlite3').verbose();

const DB_PATH = path.join(__dirname, 'todo-reminder.db');
const ENV_FILE = path.join(__dirname, '.env.local');

console.log('🚀 ローカル開発環境セットアップを開始します\n');

// ステップ 1: SQLite DB 作成
console.log('📦 ステップ 1: SQLite DB 作成中...');

const db = new sqlite3.Database(DB_PATH, (err) => {
  if (err) {
    console.error('❌ DB 作成エラー:', err);
    process.exit(1);
  }
  console.log('✅ DB を作成しました:', DB_PATH);
});

// ステップ 2: テーブル作成
console.log('\n📋 ステップ 2: テーブル作成中...');

const createTableSQL = `
CREATE TABLE IF NOT EXISTS tasks (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  due_date DATE NOT NULL,
  email TEXT NOT NULL,
  participants TEXT,
  status TEXT DEFAULT 'pending',
  board_id TEXT,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS email_history (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  task_id INTEGER,
  email TEXT NOT NULL,
  subject TEXT,
  body TEXT,
  status TEXT DEFAULT 'sent',
  error_message TEXT,
  sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS user_settings (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  user_email TEXT UNIQUE,
  reminder_days_before INTEGER DEFAULT 3,
  email_time TIME DEFAULT '09:00',
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
`;

db.exec(createTableSQL, (err) => {
  if (err) {
    console.error('❌ テーブル作成エラー:', err);
    process.exit(1);
  }
  console.log('✅ テーブルを作成しました');

  // ステップ 3: テストデータ挿入
  insertTestData();
});

function insertTestData() {
  console.log('\n📊 ステップ 3: テストデータ挿入中...');

  const today = new Date();
  const tomorrow = new Date(today.getTime() + 1 * 24 * 60 * 60 * 1000);
  const in3days = new Date(today.getTime() + 3 * 24 * 60 * 60 * 1000);

  const testTasks = [
    {
      name: 'プロジェクト企画書提出',
      due_date: tomorrow.toISOString().split('T')[0],
      email: 'test@example.com',
      participants: '田中太郎,佐藤次郎'
    },
    {
      name: 'クライアント打ち合わせ',
      due_date: in3days.toISOString().split('T')[0],
      email: 'test@example.com',
      participants: '山田花子'
    },
    {
      name: 'Q3 予算申請',
      due_date: today.toISOString().split('T')[0],
      email: 'test@example.com',
      participants: '全員'
    }
  ];

  db.serialize(() => {
    const stmt = db.prepare(`
      INSERT INTO tasks (name, due_date, email, participants, status)
      VALUES (?, ?, ?, ?, 'pending')
    `);

    testTasks.forEach(task => {
      stmt.run(task.name, task.due_date, task.email, task.participants);
    });

    stmt.finalize(() => {
      console.log(`✅ ${testTasks.length} 件のテストデータを挿入しました`);
      createEnvFile();
    });
  });
}

function createEnvFile() {
  console.log('\n⚙️  ステップ 4: 環境変数ファイル作成中...');

  const envContent = `# ローカル開発用環境変数

# Supabase（本番環境）
SUPABASE_URL=https://xxxxx.supabase.co
SUPABASE_SERVICE_KEY=eyJhbGc...

# SendGrid
SENDGRID_API_KEY=SG.xxxxx...
SENDGRID_FROM_EMAIL=noreply@todoreminder.app

# Claude AI
ANTHROPIC_API_KEY=sk-ant-xxxxx...

# ローカル DB
DATABASE_URL=sqlite:///./todo-reminder.db

# 開発モード
NODE_ENV=development
`;

  fs.writeFileSync(ENV_FILE, envContent);
  console.log('✅ .env.local ファイルを作成しました');
  console.log('   API キーを入力してください:', ENV_FILE);

  // ステップ 5: 完了メッセージ
  showCompletionMessage();
}

function showCompletionMessage() {
  console.log('\n' + '='.repeat(60));
  console.log('✅ ローカル開発環境のセットアップが完了しました！\n');

  console.log('📋 テストデータ:');
  db.all('SELECT id, name, due_date, email FROM tasks', (err, rows) => {
    if (rows) {
      rows.forEach(row => {
        console.log(`  - [${row.id}] ${row.name} (期日: ${row.due_date})`);
      });
    }

    console.log('\n🔧 次のステップ:');
    console.log('  1. .env.local ファイルに API キーを入力');
    console.log('  2. npm install で依存パッケージをインストール');
    console.log('  3. npm run dev でローカルサーバーを起動');
    console.log('  4. http://localhost:3000/api/send-reminders でテスト\n');

    console.log('📚 ドキュメント:');
    console.log('  - SERVERLESS_SETUP.md: 本番環境セットアップ手順');
    console.log('  - api/send-reminders.js: Vercel 関数のコード');
    console.log('  - .github/workflows/daily-reminder.yml: 自動実行設定\n');

    console.log('🎯 テストデータベース:');
    console.log('  - ファイル:', DB_PATH);
    console.log('  - テーブル: tasks, email_history, user_settings\n');

    console.log('='.repeat(60) + '\n');

    db.close();
  });
}

// エラーハンドリング
process.on('unhandledRejection', (reason, promise) => {
  console.error('❌ エラーが発生しました:', reason);
  process.exit(1);
});

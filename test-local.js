#!/usr/bin/env node

/**
 * ローカル API テストスクリプト
 *
 * メール送信関数をローカルでテストします
 * 実際のメール送信は行わず、ログに出力します
 */

require('dotenv').config({ path: '.env.local' });

const sqlite3 = require('sqlite3').verbose();
const Anthropic = require('@anthropic-ai/sdk').default;

const DB_PATH = './todo-reminder.db';
let db;

console.log('🧪 ローカル API テストを開始します\n');

// DB 接続
db = new sqlite3.Database(DB_PATH, (err) => {
  if (err) {
    console.error('❌ DB 接続エラー:', err);
    process.exit(1);
  }
  console.log('✅ DB に接続しました:', DB_PATH, '\n');

  testSendReminders();
});

async function testSendReminders() {
  console.log('📧 リマインドメール送信テストを開始\n');

  try {
    // 1. タスク取得
    console.log('1️⃣ タスク取得中...');
    const tasks = await getTasks();
    console.log(`   ✅ ${tasks.length} 件のタスクを取得\n`);

    if (tasks.length === 0) {
      console.log('   ⚠️  対象タスクがありません');
      closeDB();
      return;
    }

    // 2. 各タスクでメール生成テスト
    console.log('2️⃣ メール下書き生成テスト\n');

    for (const task of tasks) {
      console.log(`   📝 タスク: ${task.name}`);
      console.log(`      期日: ${task.due_date}`);
      console.log(`      メール: ${task.email}`);
      console.log(`      参加者: ${task.participants || 'なし'}`);

      try {
        // Claude API またはモック でメール生成
        const isRealAPI = !!process.env.ANTHROPIC_API_KEY;
        const apiLabel = isRealAPI ? '🤖 Claude API' : '🎭 モック（テスト用）';

        console.log(`      ${apiLabel} でメール生成中...`);
        const emailBody = await generateEmailBody(task);

        console.log('      📧 生成されたメール本文:');
        const lines = emailBody.split('\n');
        lines.forEach(line => {
          console.log(`         ${line}`);
        });
        console.log('');

      } catch (error) {
        console.error(`      ❌ エラー: ${error.message}\n`);
      }
    }

    // 3. テスト結果
    console.log('3️⃣ テスト完了\n');
    console.log('📊 テスト結果:');
    console.log(`   - タスク数: ${tasks.length}`);
    console.log(`   - 期日範囲: 本日～3日以内`);
    console.log(`   - 送信予定メール数: ${tasks.length}`);
    console.log(`   - 推定コスト: 約 $${(tasks.length * 0.001).toFixed(3)} (Claude API)\n`);

    console.log('✅ ローカルテスト完了！\n');
    console.log('次のステップ:');
    console.log('  1. API キーを .env.local に設定');
    console.log('  2. Supabase アカウント作成');
    console.log('  3. Vercel にプロジェクトをデプロイ');
    console.log('  4. 本番環境でテスト実行\n');

  } catch (error) {
    console.error('❌ テストエラー:', error);
  } finally {
    closeDB();
  }
}

async function getTasks() {
  return new Promise((resolve, reject) => {
    const now = new Date();
    const targetDate = new Date(now.getTime() + 3 * 24 * 60 * 60 * 1000);
    const targetDateStr = targetDate.toISOString().split('T')[0];
    const nowDateStr = now.toISOString().split('T')[0];

    db.all(
      `SELECT * FROM tasks
       WHERE status = 'pending'
       AND due_date <= ?
       AND due_date >= ?`,
      [targetDateStr, nowDateStr],
      (err, rows) => {
        if (err) reject(err);
        else resolve(rows || []);
      }
    );
  });
}

async function generateEmailBody(task) {
  if (!process.env.ANTHROPIC_API_KEY) {
    // API キーなしの場合はモック応答を返す
    return generateMockEmail(task);
  }

  const anthropic = new Anthropic({
    apiKey: process.env.ANTHROPIC_API_KEY
  });

  const prompt = `
以下のタスクについて、相手に丁寧にリマインドするメールの本文を作成してください。

【タスク情報】
- タスク名: ${task.name}
- 期日: ${task.due_date}
- 参加者: ${task.participants || 'なし'}

【要件】
- 日本語で丁寧な文体
- 100〜200文字程度
- 具体的で実行可能な内容

メール本文のみを出力してください。
`;

  const message = await anthropic.messages.create({
    model: 'claude-opus-4-1-20250805',
    max_tokens: 300,
    messages: [
      {
        role: 'user',
        content: prompt
      }
    ]
  });

  return message.content[0].type === 'text' ? message.content[0].text : '';
}

function generateMockEmail(task) {
  const daysUntil = Math.ceil(
    (new Date(task.due_date) - new Date()) / (1000 * 60 * 60 * 24)
  );

  const templates = {
    overdue: `お疲れ様です。\n\n「${task.name}」の提出期限がすでに過ぎています。\n大変お手数ですが、至急のご対応をお願いいたします。\n\nよろしくお願いいたします。`,

    today: `お疲れ様です。\n\n「${task.name}」の提出期限が本日となっています。\nお忙しいところ恐れ入りますが、本日中のご完了をお願いいたします。\n\nよろしくお願いいたします。`,

    soon: `お疲れ様です。\n\n「${task.name}」の提出期限が${daysUntil}日後に迫っています。\nお手すきの際に、ご対応いただきますようお願いいたします。\n\nよろしくお願いいたします。`
  };

  if (daysUntil < 0) return templates.overdue;
  if (daysUntil === 0) return templates.today;
  return templates.soon;
}

function closeDB() {
  db.close((err) => {
    if (err) console.error('DB クローズエラー:', err);
  });
}

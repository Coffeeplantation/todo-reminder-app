/**
 * Vercel Serverless Function: メール送信機能
 *
 * 役割：
 * - Supabase から期日が近いタスクを取得
 * - Claude API でメール下書きを生成
 * - SendGrid でメール送信
 * - 送信履歴を記録
 */

import { createClient } from '@supabase/supabase-js';
import sgMail from '@sendgrid/mail';
import Anthropic from '@anthropic-ai/sdk';

// サービス初期化
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY
);

sgMail.setApiKey(process.env.SENDGRID_API_KEY);
const anthropic = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY
});

/**
 * メール送信のメイン処理
 */
export default async function handler(req, res) {
  console.log('📧 リマインドメール送信プロセス開始');

  try {
    // 1. 設定を取得（デフォルト値）
    const reminderDaysBefore = 3;
    const now = new Date();
    const targetDate = new Date(now.getTime() + reminderDaysBefore * 24 * 60 * 60 * 1000);
    const targetDateStr = targetDate.toISOString().split('T')[0];

    console.log(`⏰ 対象日付: ${targetDateStr}（${reminderDaysBefore}日以内）`);

    // 2. Supabase からタスクを取得
    const { data: tasks, error: fetchError } = await supabase
      .from('tasks')
      .select('*')
      .eq('status', 'pending')
      .lte('due_date', targetDateStr)
      .gte('due_date', now.toISOString().split('T')[0]);

    if (fetchError) {
      throw new Error(`Supabase fetch error: ${fetchError.message}`);
    }

    console.log(`📋 ${tasks?.length || 0} 件のタスクを取得`);

    if (!tasks || tasks.length === 0) {
      console.log('✅ 対象タスクなし');
      return res.status(200).json({
        success: true,
        message: 'No tasks to remind',
        sent: 0
      });
    }

    // 3. 各タスクについてメール送信
    let sentCount = 0;
    const errors = [];

    for (const task of tasks) {
      try {
        // メール下書きを生成
        const emailBody = await generateEmailBody(task);

        // メール送信
        await sendEmail(task.email, task.name, emailBody);

        // 送信履歴を記録
        await recordEmailHistory(
          task.id,
          task.email,
          `リマインド: ${task.name}`,
          emailBody,
          'sent'
        );

        sentCount++;
        console.log(`✅ メール送信完了: ${task.email}`);

      } catch (error) {
        console.error(`❌ タスク ${task.id} でエラー:`, error.message);
        errors.push({
          taskId: task.id,
          email: task.email,
          error: error.message
        });

        // エラー履歴を記録
        await recordEmailHistory(
          task.id,
          task.email,
          `リマインド: ${task.name}`,
          '',
          'failed',
          error.message
        );
      }
    }

    console.log(`📊 送信完了: ${sentCount}/${tasks.length}`);

    return res.status(200).json({
      success: true,
      sent: sentCount,
      total: tasks.length,
      errors: errors.length > 0 ? errors : undefined
    });

  } catch (error) {
    console.error('❌ リマインドメール送信エラー:', error);
    return res.status(500).json({
      success: false,
      error: error.message
    });
  }
}

/**
 * Claude API でメール本文を生成
 */
async function generateEmailBody(task) {
  const prompt = `
以下のタスクについて、相手に丁寧にリマインドするメールの本文を作成してください。

【タスク情報】
- タスク名: ${task.name}
- 期日: ${task.due_date}
- 参加者: ${task.participants?.join(', ') || 'なし'}

【要件】
- 日本語で丁寧な文体
- 100〜200文字程度
- 具体的で実行可能な内容
- ビジネスメールの形式

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

/**
 * SendGrid でメール送信
 */
async function sendEmail(to, taskName, body) {
  const msg = {
    to,
    from: process.env.SENDGRID_FROM_EMAIL || 'noreply@todoreminder.app',
    subject: `📅 リマインド: ${taskName}`,
    text: body,
    html: `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px;">
        <h2 style="color: #2c3e50;">📅 タスクリマインド</h2>
        <div style="background: #f8f9fa; padding: 15px; border-radius: 8px; margin: 20px 0;">
          <p style="line-height: 1.8; color: #333; white-space: pre-wrap;">${escapeHtml(body)}</p>
        </div>
        <p style="color: #7f8c8d; font-size: 12px;">
          このメールは Todoリマインダー から自動送信されています。
        </p>
      </div>
    `
  };

  await sgMail.send(msg);
}

/**
 * 送信履歴を Supabase に記録
 */
async function recordEmailHistory(taskId, email, subject, body, status, errorMessage = null) {
  const { error } = await supabase
    .from('email_history')
    .insert({
      task_id: taskId,
      email,
      subject,
      body,
      status,
      error_message: errorMessage,
      sent_at: new Date().toISOString()
    });

  if (error) {
    console.error('履歴記録エラー:', error);
  }
}

/**
 * HTML エスケープ
 */
function escapeHtml(text) {
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return text.replace(/[&<>"']/g, m => map[m]);
}

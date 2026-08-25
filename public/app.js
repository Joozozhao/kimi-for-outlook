// Kimi for Outlook - 主逻辑
// Office.js 加载完成后初始化

const API_BASE = 'https://api.moonshot.cn/v1';
const MODEL = 'moonshot-v1-8k';

let currentEmail = { subject: '', sender: '', body: '' };
let apiKey = '';

// ===== 初始化 =====
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loadApiKey();
    setupTabs();
    setupButtons();
    refreshEmailInfo();
  }
});

function loadApiKey() {
  apiKey = localStorage.getItem('kimi_api_key') || '';
  if (!apiKey) {
    showSettings();
  }
}

// ===== 获取邮件信息 =====
function refreshEmailInfo() {
  const item = Office.context.mailbox.item;
  if (!item) {
    document.getElementById('email-subject').textContent = '请打开一封邮件使用此插件';
    return;
  }

  document.getElementById('email-subject').textContent = item.subject || '(无主题)';

  let senderText = '';
  if (item.from) {
    senderText = `发件人: ${item.from.displayName || item.from.emailAddress}`;
  }
  document.getElementById('email-sender').textContent = senderText;

  // 获取邮件正文
  currentEmail.subject = item.subject || '';
  if (item.from) {
    currentEmail.sender = item.from.displayName || item.from.emailAddress || '';
  }

  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      currentEmail.body = result.value || '';
    }
  });
}

// ===== Tab 切换 =====
function setupTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(`tab-${btn.dataset.tab}`).classList.add('active');
    });
  });
}

// ===== 设置面板 =====
function setupButtons() {
  document.getElementById('settings-btn').addEventListener('click', showSettings);
  document.getElementById('save-key-btn').addEventListener('click', saveApiKey);
  document.getElementById('cancel-settings-btn').addEventListener('click', hideSettings);

  document.getElementById('btn-triage').addEventListener('click', doTriage);
  document.getElementById('btn-reply').addEventListener('click', doReply);
  document.getElementById('btn-meeting').addEventListener('click', doMeeting);
  document.getElementById('btn-tone').addEventListener('click', doTone);

  document.getElementById('btn-insert-reply').addEventListener('click', () => insertText('result-reply'));
  document.getElementById('btn-insert-tone').addEventListener('click', () => insertText('result-tone'));
}

function showSettings() {
  document.getElementById('api-key-input').value = apiKey;
  document.getElementById('settings-panel').classList.remove('hidden');
  document.getElementById('main-panel').classList.add('hidden');
}

function hideSettings() {
  document.getElementById('settings-panel').classList.add('hidden');
  document.getElementById('main-panel').classList.remove('hidden');
}

function saveApiKey() {
  const key = document.getElementById('api-key-input').value.trim();
  if (!key) {
    alert('请输入 API Key');
    return;
  }
  apiKey = key;
  localStorage.setItem('kimi_api_key', apiKey);
  hideSettings();
}

// ===== 调用 Kimi API =====
async function callKimi(systemPrompt, userPrompt) {
  if (!apiKey) {
    showSettings();
    throw new Error('请先设置 API Key');
  }

  showLoading(true);
  try {
    const response = await fetch(`${API_BASE}/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: MODEL,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userPrompt }
        ],
        temperature: 0.7
      })
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err.error?.message || `HTTP ${response.status}`);
    }

    const data = await response.json();
    return data.choices[0].message.content;
  } finally {
    showLoading(false);
  }
}

function showLoading(show) {
  document.getElementById('loading').classList.toggle('hidden', !show);
}

function showResult(id, html) {
  const el = document.getElementById(id);
  el.innerHTML = html;
  el.classList.add('visible');
}

function getEmailContext() {
  return `邮件主题：${currentEmail.subject}\n发件人：${currentEmail.sender}\n邮件内容：\n${currentEmail.body}`;
}

// ===== 功能 1：邮件分类 =====
async function doTriage() {
  if (!currentEmail.body) {
    showResult('result-triage', '<p style="color:#cf222e">请先打开一封邮件</p>');
    return;
  }

  const system = `你是一位专业的邮件助理。请分析以下邮件，给出结构化评估。用中文回复，格式如下：

**优先级**：紧急/高/中/低（附带颜色标签说明）
**类别**：工作/会议/通知/营销/社交
**需要回复**：是/否/可延后
**建议操作**：立即处理/稍后处理/归档/删除
**关键信息**：提取邮件中最重要的1-3个要点
**待办事项**：列出需要采取的具体行动

保持简洁专业。`;

  try {
    const result = await callKimi(system, getEmailContext());
    // 简单格式化：把 markdown 标签转成 HTML
    let html = escapeHtml(result)
      .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
      .replace(/\n/g, '<br>');
    showResult('result-triage', html);
  } catch (e) {
    showResult('result-triage', `<p style="color:#cf222e">错误：${escapeHtml(e.message)}</p>`);
  }
}

// ===== 功能 2：起草回复 =====
async function doReply() {
  if (!currentEmail.body) {
    showResult('result-reply', '<p style="color:#cf222e">请先打开一封邮件</p>');
    return;
  }

  const intent = document.getElementById('reply-intent').value.trim();
  const tone = document.getElementById('reply-tone').value;

  const toneMap = {
    formal: '正式商务语气，适合对外、上级、客户',
    friendly: '友好亲和语气，适合同事和合作伙伴',
    concise: '简洁直接，适合内部快速沟通',
    warm: '热情感谢的语气'
  };

  const system = `你是一位专业的邮件撰写助手。请根据原邮件内容和用户意图，起草一封得体的回复邮件。

语气要求：${toneMap[tone]}

要求：
1. 包含恰当的称呼和结尾
2. 保持上下文连贯，必要时引用原邮件要点
3. 语言自然流畅，避免 AI 腔
4. 直接输出邮件正文，不需要主题行
${intent ? `5. 用户特别要求：${intent}` : ''}`;

  try {
    const result = await callKimi(system, getEmailContext());
    showResult('result-reply', `<pre style="white-space:pre-wrap;font-family:inherit">${escapeHtml(result)}</pre>`);
    document.getElementById('btn-insert-reply').classList.remove('hidden');
  } catch (e) {
    showResult('result-reply', `<p style="color:#cf222e">错误：${escapeHtml(e.message)}</p>`);
  }
}

// ===== 功能 3：安排会议 =====
async function doMeeting() {
  if (!currentEmail.body) {
    showResult('result-meeting', '<p style="color:#cf222e">请先打开一封邮件</p>');
    return;
  }

  const duration = document.getElementById('meeting-duration').value;
  const timePref = document.getElementById('meeting-time').value.trim();

  const system = `你是一位高效的行政助理。请从邮件中提取会议相关信息，并生成会议安排建议。

输出格式：
**会议主题**：提取或建议的主题
**参会人员**：从邮件中提取的相关人员
**建议时长**：${duration}分钟
**建议时间**：${timePref || '根据邮件内容建议'}
**会议议程**：列出3-5个要点
**需要准备的材料**：如有
**邀请邮件草稿**：一段可以直接发送的会议邀请文字

用中文输出，保持专业简洁。`;

  try {
    const result = await callKimi(system, getEmailContext());
    let html = escapeHtml(result)
      .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
      .replace(/\n/g, '<br>');
    showResult('result-meeting', html);
  } catch (e) {
    showResult('result-meeting', `<p style="color:#cf222e">错误：${escapeHtml(e.message)}</p>`);
  }
}

// ===== 功能 4：调整语气 =====
async function doTone() {
  if (!currentEmail.body) {
    showResult('result-tone', '<p style="color:#cf222e">请先打开一封邮件</p>');
    return;
  }

  const targetTone = document.getElementById('target-tone').value;

  const toneMap = {
    formal: '正式商务语气，措辞严谨、礼貌，适合对外沟通、向上汇报或正式场合',
    friendly: '友好亲和的语气，像熟悉的同事之间交流，自然轻松但不失专业',
    concise: '极度简洁直接，去掉所有冗余词汇，只保留核心信息，适合内部快速同步',
    diplomatic: '委婉礼貌的语气，用于表达拒绝、请求延期、提出异议等敏感场景',
    urgent: '紧急但有礼貌的催促语气，传达时间紧迫性但不失职业素养'
  };

  const system = `你是一位资深的商务沟通专家。请将以下邮件内容改写为指定语气，同时保持原意不变。

目标语气：${toneMap[targetTone]}

要求：
1. 保持邮件的核心信息和目的不变
2. 只改写措辞和语气，不添加新内容
3. 输出完整的改写后邮件正文
4. 如果原邮件是回复，保持回复的上下文感`;

  try {
    const result = await callKimi(system, getEmailContext());
    showResult('result-tone', `<pre style="white-space:pre-wrap;font-family:inherit">${escapeHtml(result)}</pre>`);
    document.getElementById('btn-insert-tone').classList.remove('hidden');
  } catch (e) {
    showResult('result-tone', `<p style="color:#cf222e">错误：${escapeHtml(e.message)}</p>`);
  }
}

// ===== 插入到邮件正文 =====
function insertText(resultId) {
  const text = document.getElementById(resultId).textContent;
  if (!text) return;

  const item = Office.context.mailbox.item;
  if (!item) {
    alert('无法访问邮件正文');
    return;
  }

  // 在撰写模式下插入
  item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      // 成功
    } else {
      alert('插入失败：' + result.error.message);
    }
  });
}

// ===== 工具函数 =====
function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

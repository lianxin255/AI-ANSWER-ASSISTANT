/**
 * 题库管理页面 UI 逻辑
 */
(function () {
  // DOM 引用
  const backBtn = document.getElementById('backBtn');
  const importBtn = document.getElementById('importBtn');
  const fileInput = document.getElementById('fileInput');
  const bankList = document.getElementById('bankList');
  const emptyBanks = document.getElementById('emptyBanks');
  const bindingList = document.getElementById('bindingList');
  const emptyBindings = document.getElementById('emptyBindings');
  const bindBankSelect = document.getElementById('bindBankSelect');
  const sitePatternInput = document.getElementById('sitePattern');
  const siteNameInput = document.getElementById('siteName');
  const saveBindingBtn = document.getElementById('saveBindingBtn');
  const importModal = document.getElementById('importModal');
  const importMessage = document.getElementById('importMessage');
  const importProgress = document.getElementById('importProgress');
  const importStatus = document.getElementById('importStatus');
  const importResult = document.getElementById('importResult');
  const resultMessage = document.getElementById('resultMessage');
  const importDoneBtn = document.getElementById('importDoneBtn');

  // 预览弹窗 DOM
  const previewModal = document.getElementById('previewModal');
  const previewTitle = document.getElementById('previewTitle');
  const previewSearch = document.getElementById('previewSearch');
  const previewCount = document.getElementById('previewCount');
  const previewList = document.getElementById('previewList');
  const closePreviewBtn = document.getElementById('closePreviewBtn');

  // 预览状态
  let previewData = [];

  // Tab 切换
  document.querySelectorAll('.tab-item').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab-item').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      document.getElementById('panel' + capitalize(tab.dataset.tab)).classList.add('active');
    });
  });

  function capitalize(s) {
    return s.charAt(0).toUpperCase() + s.slice(1);
  }

  // 返回按钮
  backBtn.addEventListener('click', () => window.close());

  // ============ 题库列表 ============

  async function renderBankList() {
    const banks = await QuestionBankManager.getAllBanks();
    const entries = Object.values(banks).sort((a, b) => b.importTime - a.importTime);

    if (entries.length === 0) {
      bankList.innerHTML = '';
      emptyBanks.style.display = '';
      return;
    }
    emptyBanks.style.display = 'none';

    bankList.innerHTML = entries.map(bank => `
      <div class="bank-card" data-id="${bank.id}">
        <div class="bank-header">
          <div class="bank-info">
            <h3 class="bank-name">${escapeHtml(bank.name)}</h3>
            <div class="bank-meta">
              <span class="bank-badge count">${bank.questionCount} 题</span>
              <span class="bank-badge format">${bank.formatType || '未知'}</span>
              <span>${formatDate(bank.importTime)}</span>
            </div>
          </div>
          <div class="bank-actions">
            <button class="btn btn-secondary btn-sm" onclick="event.stopPropagation(); renameBank('${bank.id}')">重命名</button>
            <button class="btn btn-danger btn-sm" onclick="event.stopPropagation(); deleteBank('${bank.id}')">删除</button>
          </div>
        </div>
        ${bank.stats.matchCount > 0 ? `
        <div class="bank-stats">
          <div class="stat-item">
            <span class="stat-label">匹配次数</span>
            <span class="stat-value">${bank.stats.matchCount}</span>
          </div>
          <div class="stat-item">
            <span class="stat-label">最近使用</span>
            <span class="stat-value">${bank.stats.lastUsed ? formatDate(bank.stats.lastUsed) : '-'}</span>
          </div>
        </div>` : ''}
        <div class="bank-view-hint">点击查看题目</div>
      </div>
    `).join('');

    // 绑定点击事件 - 打开预览
    bankList.querySelectorAll('.bank-card').forEach(card => {
      card.addEventListener('click', () => {
        const bankId = card.dataset.id;
        openPreview(bankId);
      });
    });

    // 同步更新关联表单中的题库选择
    renderBankSelectCheckboxes(entries);
  }

  function renderBankSelectCheckboxes(banks) {
    bindBankSelect.innerHTML = banks.length === 0
      ? '<p style="font-size:12px;color:#9ca3af;">请先导入题库</p>'
      : banks.map(b => `
        <label>
          <input type="checkbox" value="${b.id}" name="bindBank">
          ${escapeHtml(b.name)} (${b.questionCount}题)
        </label>
      `).join('');
  }

  // 全局函数：重命名
  window.renameBank = async function (bankId) {
    const bank = await QuestionBankManager.getBank(bankId);
    if (!bank) return;
    const newName = prompt('输入新名称:', bank.name);
    if (newName && newName.trim()) {
      await QuestionBankManager.renameBank(bankId, newName.trim());
      renderBankList();
    }
  };

  // 全局函数：删除
  window.deleteBank = async function (bankId) {
    const bank = await QuestionBankManager.getBank(bankId);
    if (!bank) return;
    if (confirm(`确定删除题库「${bank.name}」？此操作不可恢复。`)) {
      await QuestionBankManager.deleteBank(bankId);
      renderBankList();
      renderBindingList();
    }
  };

  // ============ 题库预览 ============

  async function openPreview(bankId) {
    const bank = await QuestionBankManager.getBank(bankId);
    if (!bank) return;

    previewTitle.textContent = bank.name;
    previewSearch.value = '';
    previewModal.style.display = '';

    previewData = await QuestionBankManager.getBankData(bankId);
    renderPreviewList(previewData);
  }

  function renderPreviewList(questions) {
    previewCount.textContent = `${questions.length} 题`;

    if (questions.length === 0) {
      previewList.innerHTML = '<div class="preview-empty">无匹配题目</div>';
      return;
    }

    // 只渲染前 100 题以避免性能问题
    const showQuestions = questions.slice(0, 100);
    const typeLabels = { single: '单选', multiple: '多选', judge: '判断', fill: '填空' };

    previewList.innerHTML = showQuestions.map((q, i) => {
      const typeClass = q.type || 'single';
      const typeLabel = typeLabels[typeClass] || typeClass;
      const optionsHtml = (q.options || []).map(opt =>
        `<span>${escapeHtml(opt.label)}. ${escapeHtml(opt.text)}</span>`
      ).join('');

      return `
        <div class="preview-item">
          <div class="preview-item-header">
            <span class="preview-item-num">#${i + 1}</span>
            <span class="preview-type-badge ${typeClass}">${typeLabel}</span>
          </div>
          <div class="preview-stem">${escapeHtml(q.stem)}</div>
          ${optionsHtml ? `<div class="preview-options">${optionsHtml}</div>` : ''}
          <div class="preview-answer">答案: ${escapeHtml(q.answer || '')}</div>
        </div>
      `;
    }).join('');

    if (questions.length > 100) {
      previewList.innerHTML += `<div class="preview-empty">仅显示前 100 题，共 ${questions.length} 题</div>`;
    }
  }

  // 搜索功能
  let searchTimer = null;
  previewSearch.addEventListener('input', () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      const keyword = previewSearch.value.trim().toLowerCase();
      if (!keyword) {
        renderPreviewList(previewData);
        return;
      }
      const filtered = previewData.filter(q =>
        (q.stem || '').toLowerCase().includes(keyword) ||
        (q.answer || '').toLowerCase().includes(keyword) ||
        (q.options || []).some(opt => (opt.text || '').toLowerCase().includes(keyword))
      );
      renderPreviewList(filtered);
    }, 200);
  });

  // 关闭预览
  closePreviewBtn.addEventListener('click', () => {
    previewModal.style.display = 'none';
    previewData = [];
  });

  previewModal.addEventListener('click', (e) => {
    if (e.target === previewModal) {
      previewModal.style.display = 'none';
      previewData = [];
    }
  });

  // ============ 导入题库 ============

  importBtn.addEventListener('click', () => fileInput.click());

  fileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    fileInput.value = ''; // 重置以便重复选择同一文件

    // 显示导入弹窗
    importModal.style.display = '';
    importStatus.style.display = '';
    importResult.style.display = 'none';
    importMessage.textContent = `正在解析：${file.name}`;
    importProgress.style.width = '30%';

    try {
      const buffer = await file.arrayBuffer();
      importProgress.style.width = '60%';

      const parsed = await XlsxParser.parse(buffer, {
        onAIFallback: async ({ headers, sampleRows }) => {
          importMessage.textContent = '格式未知，正在请求 AI 解析...';
          return await requestAIMapping(headers, sampleRows);
        }
      });
      const { questions } = parsed;
      const formatType = parsed.formatType || parsed.format || '未知';

      importProgress.style.width = '90%';

      if (questions.length === 0) {
        throw new Error('解析结果为空，未找到有效题目');
      }

      // 存储
      const bankId = await QuestionBankManager.importBank(file.name, questions, formatType);
      importProgress.style.width = '100%';

      // 显示结果
      importStatus.style.display = 'none';
      importResult.style.display = '';
      resultMessage.textContent = `导入成功！共解析 ${questions.length} 道题目（${formatType}格式）`;

    } catch (err) {
      importStatus.style.display = 'none';
      importResult.style.display = '';
      resultMessage.textContent = `导入失败：${err.message}`;
    }
  });

  importDoneBtn.addEventListener('click', () => {
    importModal.style.display = 'none';
    renderBankList();
  });

  /** 请求 AI 解析列映射 */
  async function requestAIMapping(headers, sampleRows) {
    try {
      const prompt = `我有一个题库Excel文件，请分析其列结构并返回JSON映射。

表头: ${JSON.stringify(headers)}
前5行样本: ${JSON.stringify(sampleRows.slice(0, 3))}

请返回JSON格式的列映射（只返回JSON，不要其他文字）：
{
  "stem": "题干所在列的表头名",
  "type": "题型所在列的表头名（如有）",
  "answer": "答案所在列的表头名",
  "options": "选项合并列的表头名（如果选项在一列中）",
  "optionStart": "选项起始列的表头名（如果选项分多列）",
  "optionEnd": "选项结束列的表头名（如果选项分多列）"
}
注意：options和optionStart/optionEnd二选一，没有的字段不要包含。`;

      const response = await chrome.runtime.sendMessage({
        action: 'callAI',
        config: await getActiveModelConfig(),
        prompt,
        mode: 'mapping'
      });

      if (response && response.success && response.data) {
        const raw = response.data;
        const text = typeof raw === 'string' ? raw : (raw.text || JSON.stringify(raw));
        const jsonStr = text.replace(/```json?\s*|\s*```/g, '').trim();
        return JSON.parse(jsonStr);
      }
    } catch (err) {
      console.error('[题库] AI 映射解析失败:', err);
    }
    return null;
  }

  /** 获取活跃模型配置 */
  async function getActiveModelConfig() {
    const result = await chrome.storage.sync.get(['aiModels', 'activeModelId']);
    const models = result.aiModels || [];
    const activeId = result.activeModelId;
    const model = models.find(m => m.id === activeId) || models[0];
    if (!model) return { baseUrl: '', apiKey: '', model: '' };
    return { baseUrl: model.baseUrl, apiKey: model.apiKey, model: model.model };
  }

  // ============ 站点关联 ============

  async function renderBindingList() {
    const bindings = await QuestionBankManager.getAllBindings();
    const banks = await QuestionBankManager.getAllBanks();
    const entries = Object.entries(bindings);

    if (entries.length === 0) {
      bindingList.innerHTML = '';
      emptyBindings.style.display = '';
      return;
    }
    emptyBindings.style.display = 'none';

    bindingList.innerHTML = entries.map(([pattern, binding]) => {
      const bankNames = binding.boundBankIds
        .map(id => banks[id]?.name || id)
        .map(name => `<span class="binding-bank-tag">${escapeHtml(name)}</span>`)
        .join('');

      return `
        <div class="binding-card">
          <div class="binding-header">
            <span class="binding-site">
              ${escapeHtml(pattern)}
              ${binding.siteName ? `<span class="binding-site-name">${escapeHtml(binding.siteName)}</span>` : ''}
            </span>
            <button class="btn btn-danger btn-sm" onclick="removeBinding('${escapeHtml(pattern)}')">删除</button>
          </div>
          <div class="binding-banks">${bankNames || '<span style="font-size:12px;color:#9ca3af;">无关联题库</span>'}</div>
        </div>
      `;
    }).join('');
  }

  saveBindingBtn.addEventListener('click', async () => {
    const pattern = sitePatternInput.value.trim();
    if (!pattern) {
      alert('请输入站点域名');
      return;
    }

    const selectedBankIds = Array.from(
      document.querySelectorAll('input[name="bindBank"]:checked')
    ).map(cb => cb.value);

    if (selectedBankIds.length === 0) {
      alert('请至少选择一个题库');
      return;
    }

    const siteName = siteNameInput.value.trim() || pattern;
    await QuestionBankManager.setBinding(pattern, siteName, selectedBankIds);

    sitePatternInput.value = '';
    siteNameInput.value = '';
    document.querySelectorAll('input[name="bindBank"]').forEach(cb => cb.checked = false);
    renderBindingList();
  });

  window.removeBinding = async function (pattern) {
    if (confirm(`确定删除站点「${pattern}」的关联？`)) {
      await QuestionBankManager.removeBinding(pattern);
      renderBindingList();
    }
  };

  // ============ 工具函数 ============

  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  function formatDate(ts) {
    if (!ts) return '-';
    const d = new Date(ts);
    return `${d.getMonth() + 1}/${d.getDate()} ${d.getHours()}:${String(d.getMinutes()).padStart(2, '0')}`;
  }

  // 初始化
  renderBankList();
  renderBindingList();
})();

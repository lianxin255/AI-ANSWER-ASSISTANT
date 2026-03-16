/**
 * 题库管理器
 * 负责题库 CRUD 和站点关联管理，使用 chrome.storage.local 存储
 */
class QuestionBankManager {
  /** 获取所有题库元数据 */
  static async getAllBanks() {
    const result = await chrome.storage.local.get('questionBanks');
    return result.questionBanks || {};
  }

  /** 获取单个题库元数据 */
  static async getBank(bankId) {
    const banks = await QuestionBankManager.getAllBanks();
    return banks[bankId] || null;
  }

  /** 获取题库题目数据 */
  static async getBankData(bankId) {
    const key = `bank_data_${bankId}`;
    const result = await chrome.storage.local.get(key);
    return result[key] || [];
  }

  /**
   * 导入题库
   * @param {string} fileName - 文件名
   * @param {Array} questions - 标准化题目数组
   * @param {string} formatType - 格式类型
   * @returns {string} bankId
   */
  static async importBank(fileName, questions, formatType) {
    const bankId = 'bank_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 6);

    const meta = {
      id: bankId,
      name: fileName.replace(/\.(xlsx?|csv)$/i, ''),
      fileName,
      questionCount: questions.length,
      formatType,
      importTime: Date.now(),
      stats: { matchCount: 0, lastUsed: null }
    };

    // 分别存储元数据和题目数据
    const banks = await QuestionBankManager.getAllBanks();
    banks[bankId] = meta;

    await chrome.storage.local.set({
      questionBanks: banks,
      [`bank_data_${bankId}`]: questions
    });

    return bankId;
  }

  /** 删除题库 */
  static async deleteBank(bankId) {
    const banks = await QuestionBankManager.getAllBanks();
    delete banks[bankId];

    // 同时清理关联中引用此题库的记录
    const bindings = await QuestionBankManager.getAllBindings();
    for (const [pattern, binding] of Object.entries(bindings)) {
      binding.boundBankIds = binding.boundBankIds.filter(id => id !== bankId);
      binding.activeBankIds = binding.activeBankIds.filter(id => id !== bankId);
      if (binding.boundBankIds.length === 0) delete bindings[pattern];
    }

    await chrome.storage.local.set({ questionBanks: banks, bankSiteBindings: bindings });
    await chrome.storage.local.remove(`bank_data_${bankId}`);
  }

  /** 重命名题库 */
  static async renameBank(bankId, newName) {
    const banks = await QuestionBankManager.getAllBanks();
    if (banks[bankId]) {
      banks[bankId].name = newName;
      await chrome.storage.local.set({ questionBanks: banks });
    }
  }

  /** 更新题库匹配统计 */
  static async updateBankStats(bankId, matchCount) {
    const banks = await QuestionBankManager.getAllBanks();
    if (banks[bankId]) {
      banks[bankId].stats.matchCount += matchCount;
      banks[bankId].stats.lastUsed = Date.now();
      await chrome.storage.local.set({ questionBanks: banks });
    }
  }

  // ============ 站点关联管理 ============

  /** 获取所有站点关联 */
  static async getAllBindings() {
    const result = await chrome.storage.local.get('bankSiteBindings');
    return result.bankSiteBindings || {};
  }

  /** 根据 URL 查找匹配的站点关联 */
  static async findBindingByUrl(url) {
    const bindings = await QuestionBankManager.getAllBindings();
    const hostname = new URL(url).hostname;

    // 精确匹配域名
    for (const [pattern, binding] of Object.entries(bindings)) {
      if (hostname === pattern || hostname.endsWith('.' + pattern)) {
        return binding;
      }
    }
    // 通配符匹配
    for (const [pattern, binding] of Object.entries(bindings)) {
      if (pattern.startsWith('*.')) {
        const base = pattern.slice(2);
        if (hostname === base || hostname.endsWith('.' + base)) return binding;
      }
    }
    return null;
  }

  /**
   * 设置站点关联
   * @param {string} sitePattern - 站点域名模式
   * @param {string} siteName - 站点显示名称
   * @param {string[]} bankIds - 关联的题库 ID 列表
   */
  static async setBinding(sitePattern, siteName, bankIds) {
    const bindings = await QuestionBankManager.getAllBindings();
    if (bankIds.length === 0) {
      delete bindings[sitePattern];
    } else {
      bindings[sitePattern] = {
        siteName,
        boundBankIds: bankIds,
        activeBankIds: bankIds // 默认全部激活
      };
    }
    await chrome.storage.local.set({ bankSiteBindings: bindings });
  }

  /** 更新激活的题库 */
  static async setActiveBanks(sitePattern, activeBankIds) {
    const bindings = await QuestionBankManager.getAllBindings();
    if (bindings[sitePattern]) {
      bindings[sitePattern].activeBankIds = activeBankIds;
      await chrome.storage.local.set({ bankSiteBindings: bindings });
    }
  }

  /** 删除站点关联 */
  static async removeBinding(sitePattern) {
    const bindings = await QuestionBankManager.getAllBindings();
    delete bindings[sitePattern];
    await chrome.storage.local.set({ bankSiteBindings: bindings });
  }
}

window.QuestionBankManager = QuestionBankManager;

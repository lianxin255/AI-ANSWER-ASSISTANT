/**
 * 题库匹配引擎
 * 三级匹配：精确 HashMap → 包含匹配 → 模糊 Jaccard
 */
class QuestionBankMatcher {
  constructor() {
    this._questions = [];       // 所有已加载题目
    this._exactMap = new Map(); // 归一化题干 → 题目（精确匹配索引）
    this._loaded = false;
    this._bankIds = [];
    this._matchStats = {};      // bankId → 匹配次数
  }

  /** 是否已加载题库 */
  isLoaded() {
    return this._loaded && this._questions.length > 0;
  }

  /** 已加载题目数 */
  get questionCount() {
    return this._questions.length;
  }

  /**
   * 加载题库数据
   * @param {string[]} bankIds - 要加载的题库 ID 列表
   */
  async loadBanks(bankIds) {
    this._questions = [];
    this._exactMap.clear();
    this._bankIds = bankIds;
    this._matchStats = {};

    for (const bankId of bankIds) {
      this._matchStats[bankId] = 0;
      const data = await QuestionBankManager.getBankData(bankId);
      for (const q of data) {
        // 确保 stemNormalized 已计算并缓存
        if (!q.stemNormalized) {
          q.stemNormalized = XlsxParser.normalize(q.stem);
        }
        this._questions.push({ ...q, bankId });
        // 建立精确匹配索引
        if (!this._exactMap.has(q.stemNormalized)) {
          this._exactMap.set(q.stemNormalized, { ...q, bankId });
        }
      }
    }

    this._loaded = this._questions.length > 0;
    console.log(`[题库匹配] 已加载 ${this._questions.length} 题，来自 ${bankIds.length} 个题库`);
  }

  /** 卸载题库 */
  unload() {
    this._questions = [];
    this._exactMap.clear();
    this._loaded = false;
    this._bankIds = [];
    this._matchStats = {};
  }

  /** 获取匹配统计 */
  getMatchStats() {
    return { ...this._matchStats };
  }

  /**
   * 匹配题目
   * @param {Object} question - 扫描到的题目 { text, options }
   * @returns {{ question: Object, level: number, score: number } | null}
   */
  match(question) {
    const stem = question.text || question.stem || '';
    if (!stem) return null;

    const normalized = XlsxParser.normalize(stem);

    // 第 1 级：精确匹配 O(1)
    const exact = this._exactMap.get(normalized);
    if (exact) {
      this._recordMatch(exact.bankId);
      return { question: exact, level: 1, score: 1.0 };
    }

    // 第 2 级：包含匹配 O(n)
    const containResult = this._containMatch(normalized);
    if (containResult) {
      this._recordMatch(containResult.question.bankId);
      return containResult;
    }

    // 第 3 级：模糊匹配 O(k*m)
    const fuzzyResult = this._fuzzyMatch(normalized, question.options);
    if (fuzzyResult) {
      this._recordMatch(fuzzyResult.question.bankId);
      return fuzzyResult;
    }

    return null;
  }

  /** 记录匹配统计 */
  _recordMatch(bankId) {
    if (bankId && this._matchStats[bankId] !== undefined) {
      this._matchStats[bankId]++;
    }
  }

  /** 包含匹配：题干互相包含 */
  _containMatch(normalized) {
    if (normalized.length < 5) return null;

    for (const q of this._questions) {
      const qNorm = q.stemNormalized;
      if (qNorm.length < 5) continue;

      if (normalized.includes(qNorm) || qNorm.includes(normalized)) {
        const lenRatio = Math.min(normalized.length, qNorm.length) / Math.max(normalized.length, qNorm.length);
        if (lenRatio > 0.5) {
          return { question: q, level: 2, score: lenRatio };
        }
      }
    }
    return null;
  }

  /** 模糊匹配：Jaccard bigram 相似度 + 选项辅助验证 */
  _fuzzyMatch(normalized, questionOptions) {
    if (normalized.length < 8) return null;

    const queryBigrams = QuestionBankMatcher._bigrams(normalized);
    let bestMatch = null;
    let bestScore = 0;

    // 限制候选范围：通过长度过滤
    const candidates = this._questions.filter(q => {
      const lenRatio = Math.min(normalized.length, q.stemNormalized.length) / Math.max(normalized.length, q.stemNormalized.length);
      return lenRatio > 0.4;
    });

    for (const q of candidates) {
      const qBigrams = QuestionBankMatcher._bigrams(q.stemNormalized);
      let score = QuestionBankMatcher._jaccardSimilarity(queryBigrams, qBigrams);

      // 选项辅助验证
      if (questionOptions && questionOptions.length > 0 && q.options && q.options.length > 0) {
        const optionBonus = QuestionBankMatcher._optionSimilarity(questionOptions, q.options);
        score = score * 0.7 + optionBonus * 0.3;
      }

      if (score > bestScore) {
        bestScore = score;
        bestMatch = q;
      }
    }

    if (bestScore >= 0.6 && bestMatch) {
      return { question: bestMatch, level: 3, score: bestScore };
    }
    return null;
  }

  /** 提取 bigram 集合 */
  static _bigrams(text) {
    const set = new Set();
    for (let i = 0; i < text.length - 1; i++) {
      set.add(text.slice(i, i + 2));
    }
    return set;
  }

  /** Jaccard 相似度 */
  static _jaccardSimilarity(setA, setB) {
    if (setA.size === 0 && setB.size === 0) return 0;
    let intersection = 0;
    for (const item of setA) {
      if (setB.has(item)) intersection++;
    }
    return intersection / (setA.size + setB.size - intersection);
  }

  /** 选项文本相似度（匹配选项数 / 总选项数） */
  static _optionSimilarity(scannedOptions, bankOptions) {
    const normalize = (text) => (text || '').replace(/[^a-z0-9\u4e00-\u9fff]/gi, '').toLowerCase();

    const bankTexts = bankOptions.map(o => normalize(o.text));
    let matchCount = 0;

    for (const opt of scannedOptions) {
      const optText = normalize(opt.text || opt);
      if (bankTexts.some(bt => bt.includes(optText) || optText.includes(bt))) {
        matchCount++;
      }
    }

    return matchCount / Math.max(scannedOptions.length, bankOptions.length);
  }
}

// 全局单例
window.questionBankMatcher = new QuestionBankMatcher();

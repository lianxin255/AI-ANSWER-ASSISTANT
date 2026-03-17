class XlsxParser {
  static async parse(buffer, options = {}) {
    const workbook = XLSX.read(buffer, { type: 'array' });
    if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) {
      throw new Error('Excel 文件为空或无法读取');
    }

    const sheetCandidate = this._pickBestSheet(workbook, options);
    if (!sheetCandidate) {
      throw new Error('未找到有效题库工作表');
    }

    const { sheetName, rows } = sheetCandidate;
    if (!rows.length) {
      throw new Error(`工作表 ${sheetName} 中没有可用数据`);
    }

    let result = this._tryStructuredFormats(rows, { sheetName, ...options });
    if (result && result.questions && result.questions.length) {
      return {
        ...result,
        sheetName,
        total: result.questions.length,
      };
    }

    if (typeof options.onAIFallback === 'function') {
      const payload = {
        sheetName,
        headers: rows[0] || [],
        sampleRows: rows.slice(1, 6),
      };

      try {
        const mapping = await options.onAIFallback(payload);
        if (mapping) {
          result = this._parseWithMapping(rows, mapping, { sheetName, ...options });
          if (result && result.questions && result.questions.length) {
            return {
              ...result,
              sheetName,
              total: result.questions.length,
            };
          }
        }
      } catch (_e) {
        // AI 映射失败则继续报错给上层处理
      }
    }

    throw new Error(`无法识别题库格式：${sheetName}`);
  }

  static normalize(text) {
    return String(text || '')
      .replace(/^\s*\d+[\.\-、．\)]\s*/, '')
      .replace(/[“”"'‘’【】\[\]（）()]/g, '')
      .replace(/\s+/g, '')
      .replace(/[，,。.!！?？:：;；、]/g, '')
      .trim()
      .toLowerCase();
  }

  static hash(text) {
    const s = this.normalize(text);
    let h = 0;
    for (let i = 0; i < s.length; i++) {
      h = (h << 5) - h + s.charCodeAt(i);
      h |= 0;
    }
    return String(h);
  }

  static _pickBestSheet(workbook) {
    let best = null;

    for (const sheetName of workbook.SheetNames) {
      const ws = workbook.Sheets[sheetName];
      if (!ws) continue;

      const rows = XLSX.utils.sheet_to_json(ws, {
        header: 1,
        raw: false,
        defval: '',
        blankrows: false,
      });

      const cleaned = rows
        .map((row) => (Array.isArray(row) ? row.map((v) => String(v ?? '').trim()) : []))
        .filter((row) => row.some((cell) => String(cell || '').trim()));

      if (cleaned.length < 2) continue;

      const score = this._scoreSheet(cleaned);
      if (!best || score > best.score) {
        best = { sheetName, rows: cleaned, score };
      }
    }

    return best;
  }

  static _scoreSheet(rows) {
    const headers = (rows[0] || []).map((h) => this._normHeader(h));
    const sampleRows = rows.slice(1, 8);

    let score = 0;

    const stemIdx = this._findStemIndex(headers);
    const answerIdx = this._findHeaderIndex(headers, ['答案', '正确答案', '标准答案', '参考答案']);
    const typeIdx = this._findHeaderIndex(headers, ['题型', '类型']);
    const parseIdx = this._findHeaderIndex(headers, ['解析', '答案解析', '说明', '备注']);

    if (stemIdx >= 0) score += 5;
    if (answerIdx >= 0) score += 5;
    if (typeIdx >= 0) score += 2;
    if (parseIdx >= 0) score += 1;

    const optionHeaders = ['a', 'b', 'c', 'd', 'e', 'f'].filter((k) => headers.includes(k));
    score += optionHeaders.length * 2;

    const avgCols =
      sampleRows.length > 0
        ? sampleRows.reduce((sum, row) => sum + row.filter(Boolean).length, 0) / sampleRows.length
        : 0;

    if (avgCols >= 3) score += 2;
    if (rows.length >= 10) score += 2;

    return score;
  }

  static _tryStructuredFormats(rows, options = {}) {
    const rawHeaders = rows[0].map((h) => String(h || '').trim());
    const headers = rawHeaders.map((h) => this._normHeader(h));
    const dataRows = rows.slice(1).filter((r) => r.some((x) => String(x || '').trim()));
    if (!dataRows.length) return null;

    const format2 = this._tryFormatMultiColumn(headers, dataRows, rawHeaders, options);
    const format1 = this._tryFormatMergedOptions(headers, dataRows, options);

    if (format2 && format1) {
      return format2.score >= format1.score ? format2 : format1;
    }

    return format2 || format1 || null;
  }

  static _tryFormatMultiColumn(headers, rows, rawHeaders = []) {
    const stemIdx = this._findStemIndex(headers);
    const answerIdx = this._findHeaderIndex(headers, ['答案', '正确答案', '标准答案', '参考答案']);
    const typeIdx = this._findHeaderIndex(headers, ['题型', '类型']);
    const analysisIdx = this._findHeaderIndex(headers, ['解析', '答案解析', '说明', '备注']);
    const scoreIdx = this._findHeaderIndex(headers, ['分值', '分数']);

    const optionIndices = {};
    for (const key of ['a', 'b', 'c', 'd', 'e', 'f']) {
      const idx = this._findHeaderIndex(headers, [key, `选项${key}`, `${key}选项`, `${key}项`]);
      if (idx >= 0) optionIndices[key.toUpperCase()] = idx;
    }

    // 兼容 "选项1(A)" / "选项2(B)" 这类表头
    if (Object.keys(optionIndices).length < 2 && rawHeaders.length) {
      rawHeaders.forEach((h, idx) => {
        const text = String(h || '').trim();
        if (!text) return;
        let m = text.match(/选项\s*\d*\s*[\(\[]?([A-F])[\)\]]?/i);
        if (!m) {
          m = text.match(/[\(\[]\s*([A-F])\s*[\)\]]/i);
        }
        if (m) {
          const label = m[1].toUpperCase();
          if (!optionIndices[label]) optionIndices[label] = idx;
        }
      });
    }

    if (stemIdx < 0 || Object.keys(optionIndices).length < 2) {
      return null;
    }

    const questions = [];
    let hit = 0;

    for (const row of rows) {
      const stem = String(row[stemIdx] || '').trim();
      if (!stem) continue;

      const options = [];
      for (const [label, idx] of Object.entries(optionIndices)) {
        const val = String(row[idx] || '').trim();
        if (val) {
          options.push({ label, text: val });
        }
      }

      const rawAnswer = answerIdx >= 0 ? row[answerIdx] : '';
      let type = this._normalizeType(typeIdx >= 0 ? row[typeIdx] : '');
      if (!type) {
        type = this._inferTypeFromRow({ rawAnswer, options });
      }

      const answer = this._parseAnswer(rawAnswer, type);
      const analysis = analysisIdx >= 0 ? String(row[analysisIdx] || '').trim() : '';
      const score = scoreIdx >= 0 ? this._toNumber(row[scoreIdx]) : null;

      questions.push({
        type,
        stem,
        stemNormalized: this.normalize(stem),
        hash: this.hash(stem),
        options,
        answer,
        analysis,
        score,
      });

      if (answer || options.length >= 2) hit++;
    }

    if (!questions.length) return null;

    return {
      format: 'multi-column-options',
      score: hit + 20,
      questions,
    };
  }

  static _tryFormatMergedOptions(headers, rows) {
    const stemIdx = this._findHeaderIndex(headers, ['题干', '题目', '题面', '标题', '试题', '内容']);
    const typeIdx = this._findHeaderIndex(headers, ['题型', '类型']);
    const answerIdx =
      this._findHeaderIndex(headers, ['答案', '正确答案', '标准答案', '参考答案']) ??
      -1;

    let mergedOptionIdx = this._findHeaderIndex(headers, ['选项', '备选项', '候选项', '答案选项']);

    const sampleRows = rows.slice(0, 8);

    if (mergedOptionIdx < 0) {
      let bestIdx = -1;
      let bestScore = -1;
      for (let i = 0; i < headers.length; i++) {
        if (i === stemIdx || i === typeIdx || i === answerIdx) continue;

        let score = 0;
        for (const row of sampleRows) {
          const val = String(row[i] || '').trim();
          if (this._looksLikeMergedOptions(val)) score++;
        }
        if (score > bestScore) {
          bestScore = score;
          bestIdx = i;
        }
      }
      if (bestScore >= 2) mergedOptionIdx = bestIdx;
    }

    if (stemIdx < 0 || mergedOptionIdx < 0) return null;

    let inferredAnswerIdx = answerIdx;
    if (inferredAnswerIdx < 0) {
      inferredAnswerIdx = this._guessAnswerColumn(headers, rows, [stemIdx, typeIdx, mergedOptionIdx]);
    }

    const analysisIdx = this._findHeaderIndex(headers, ['解析', '答案解析', '说明', '备注']);
    const scoreIdx = this._findHeaderIndex(headers, ['分值', '分数']);

    const questions = [];
    let hit = 0;

    for (const row of rows) {
      const stem = String(row[stemIdx] || '').trim();
      if (!stem) continue;

      const mergedOptions = String(row[mergedOptionIdx] || '').trim();
      const options = this._parseMergedOptions(mergedOptions);
      const rawAnswer = inferredAnswerIdx >= 0 ? row[inferredAnswerIdx] : '';

      let type = this._normalizeType(typeIdx >= 0 ? row[typeIdx] : '');
      if (!type) {
        type = this._inferTypeFromRow({ rawAnswer, options });
      }

      const answer = this._parseAnswer(rawAnswer, type);
      const analysis = analysisIdx >= 0 ? String(row[analysisIdx] || '').trim() : '';
      const score = scoreIdx >= 0 ? this._toNumber(row[scoreIdx]) : null;

      questions.push({
        type,
        stem,
        stemNormalized: this.normalize(stem),
        hash: this.hash(stem),
        options,
        answer,
        analysis,
        score,
      });

      if (answer || options.length >= 2) hit++;
    }

    if (!questions.length) return null;

    return {
      format: 'merged-options',
      score: hit + 10,
      questions,
    };
  }

  static _parseWithMapping(rows, mapping) {
    const headers = rows[0].map((h) => String(h || '').trim());
    const dataRows = rows.slice(1).filter((r) => r.some((x) => String(x || '').trim()));

    const getIdx = (nameOrList) => {
      if (!nameOrList) return -1;
      const wanted = Array.isArray(nameOrList) ? nameOrList : [nameOrList];

      for (const want of wanted) {
        const exact = headers.findIndex((h) => h === want);
        if (exact >= 0) return exact;
      }

      const normHeaders = headers.map((h) => this._normHeader(h));
      for (const want of wanted) {
        const n = this._normHeader(want);
        const exactNorm = normHeaders.findIndex((h) => h === n);
        if (exactNorm >= 0) return exactNorm;
      }

      for (const want of wanted) {
        const n = this._normHeader(want);
        const fuzzy = normHeaders.findIndex((h) => h.includes(n) || n.includes(h));
        if (fuzzy >= 0) return fuzzy;
      }

      return -1;
    };

    const stemIdx = getIdx(mapping.stem || ['题干', '题目']);
    const answerIdx = getIdx(mapping.answer || ['答案']);
    const typeIdx = getIdx(mapping.type || ['题型']);
    const analysisIdx = getIdx(mapping.analysis || ['解析']);
    const scoreIdx = getIdx(mapping.score || ['分值']);
    const mergedOptionIdx = getIdx(mapping.options || ['选项']);

    const optionIndices = {};
    for (const key of ['A', 'B', 'C', 'D', 'E', 'F']) {
      const idx = getIdx(mapping[key] || [key, `选项${key.toLowerCase()}`, `${key}选项`]);
      if (idx >= 0) optionIndices[key] = idx;
    }

    if (stemIdx < 0) {
      throw new Error('AI 映射结果缺少题干列');
    }

    const questions = [];
    for (const row of dataRows) {
      const stem = String(row[stemIdx] || '').trim();
      if (!stem) continue;

      let options = [];
      for (const [label, idx] of Object.entries(optionIndices)) {
        const val = String(row[idx] || '').trim();
        if (val) options.push({ label, text: val });
      }

      if (!options.length && mergedOptionIdx >= 0) {
        options = this._parseMergedOptions(String(row[mergedOptionIdx] || '').trim());
      }

      const rawAnswer = answerIdx >= 0 ? row[answerIdx] : '';
      let type = this._normalizeType(typeIdx >= 0 ? row[typeIdx] : '');
      if (!type) {
        type = this._inferTypeFromRow({ rawAnswer, options });
      }

      questions.push({
        type,
        stem,
        stemNormalized: this.normalize(stem),
        hash: this.hash(stem),
        options,
        answer: this._parseAnswer(rawAnswer, type),
        analysis: analysisIdx >= 0 ? String(row[analysisIdx] || '').trim() : '',
        score: scoreIdx >= 0 ? this._toNumber(row[scoreIdx]) : null,
      });
    }

    return {
      format: 'ai-mapping',
      score: questions.length,
      questions,
    };
  }

  static _normHeader(h) {
    return String(h || '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, '')
      .replace(/[：:（）()\[\]【】]/g, '')
      .replace(/^列/, '');
  }

  static _findHeaderIndex(headers, aliases) {
    if (!Array.isArray(headers) || !Array.isArray(aliases)) return -1;

    const normAliases = aliases.map((a) => this._normHeader(a));

    for (const alias of normAliases) {
      const exact = headers.findIndex((h) => h === alias);
      if (exact >= 0) return exact;
    }

    for (const alias of normAliases) {
      const fuzzy = headers.findIndex((h) => h.includes(alias) || alias.includes(h));
      if (fuzzy >= 0) return fuzzy;
    }

    return -1;
  }

  static _findStemIndex(headers) {
    if (!Array.isArray(headers)) return -1;

    const keywords = [
      { key: '题目标题', score: 6 },
      { key: '题目', score: 4 },
      { key: '标题', score: 3 },
      { key: '题面', score: 3 },
      { key: '试题', score: 3 },
      { key: '内容', score: 2 },
      { key: '题干', score: 2 },
      { key: '说明', score: -3 },
    ];

    let bestIdx = -1;
    let bestScore = -Infinity;

    headers.forEach((h, idx) => {
      let score = 0;
      for (const k of keywords) {
        if (h.includes(k.key)) score += k.score;
      }
      if (score > bestScore) {
        bestScore = score;
        bestIdx = idx;
      }
    });

    return bestScore > 0 ? bestIdx : -1;
  }

  static _normalizeType(raw) {
    const t = String(raw || '').trim().toLowerCase();
    if (!t) return '';

    if (/单选|single|radio/.test(t)) return 'single';
    if (/多选|multiple|multi|checkbox/.test(t)) return 'multiple';
    if (/判断|是非|truefalse|judge/.test(t)) return 'judge';
    if (/填空|简答|问答|主观|文本|text|blank/.test(t)) return 'fill';

    return '';
  }

  static _inferTypeFromRow({ rawAnswer, options }) {
    const ans = String(rawAnswer || '').trim();

    if (/^(对|错|正确|错误|是|否|√|×|true|false)$/i.test(ans)) {
      return 'judge';
    }

    const compact = ans
      .replace(/[，,；;、\s]/g, '')
      .toUpperCase();

    if (/^[A-F]{2,}$/.test(compact)) {
      return 'multiple';
    }

    if (/^[A-F]$/.test(compact)) {
      return 'single';
    }

    if (Array.isArray(options) && options.length >= 2) {
      return 'fill';
    }

    return 'fill';
  }

  static _parseAnswer(raw, type) {
    const text = String(raw || '').trim();
    if (!text) return '';

    if (type === 'judge') {
      if (/^(对|正确|是|√|true)$/i.test(text)) return true;
      if (/^(错|错误|否|×|false)$/i.test(text)) return false;
      return text;
    }

    if (type === 'single') {
      const m = text.toUpperCase().match(/[A-F]/);
      return m ? m[0] : text;
    }

    if (type === 'multiple') {
      const arr = text
        .toUpperCase()
        .replace(/[，,；;、\s]/g, '')
        .match(/[A-F]/g);
      return arr ? [...new Set(arr)] : text;
    }

    return text;
  }

  static _guessAnswerColumn(headers, rows, excludedIdx = []) {
    const sampleRows = rows.slice(0, 8);
    let bestIdx = -1;
    let bestScore = -1;

    for (let i = 0; i < headers.length; i++) {
      if (excludedIdx.includes(i)) continue;

      let score = 0;
      for (const row of sampleRows) {
        const val = String(row[i] || '').trim();
        if (!val) continue;

        if (/^(对|错|正确|错误|是|否|√|×|true|false)$/i.test(val)) score += 3;
        else if (/^[A-F]$/i.test(val)) score += 3;
        else if (/^[A-F][,，、;；\s]*[A-F]+$/i.test(val)) score += 4;
        else if (/^[A-F]{2,}$/i.test(val)) score += 4;
      }

      if (score > bestScore) {
        bestScore = score;
        bestIdx = i;
      }
    }

    return bestScore >= 4 ? bestIdx : -1;
  }

  static _looksLikeMergedOptions(text) {
    const t = String(text || '').trim();
    if (!t) return false;

    return (
      /A[\.\-:：、）)]/.test(t) ||
      /B[\.\-:：、）)]/.test(t) ||
      /\bA\b.*\bB\b/.test(t) ||
      /A[\.、:：）)]/.test(t)
    );
  }

  static _parseMergedOptions(text) {
    const raw = String(text || '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/\r/g, '\n')
      .trim();

    if (!raw) return [];

    const normalized = raw
      .replace(/[Ａ-Ｆ]/g, (m) => String.fromCharCode(m.charCodeAt(0) - 65248))
      .replace(/([A-F])[．。]/g, '$1.')
      .replace(/([A-F])[：:]/g, '$1:')
      .replace(/([A-F])[）)]/g, '$1)')
      .replace(/\u00A0/g, ' ');

    const pieces = normalized
      .replace(/([A-F][\.\-:：、\)])/g, '\n$1')
      .split(/\n+/)
      .map((s) => s.trim())
      .filter(Boolean);

    const options = [];
    for (const part of pieces) {
      const m = part.match(/^([A-F])[\.\-:：、\)]\s*(.+)$/i);
      if (m) {
        options.push({
          label: m[1].toUpperCase(),
          text: m[2].trim(),
        });
      }
    }

    if (options.length >= 2) return options;

    const fallback = [];
    const reg = /([A-F])[\.\-:：、\)]\s*([\s\S]*?)(?=(?:[A-F][\.\-:：、\)])|$)/gi;
    let match;
    while ((match = reg.exec(normalized))) {
      fallback.push({
        label: match[1].toUpperCase(),
        text: String(match[2] || '').trim(),
      });
    }

    return fallback;
  }

  static _toNumber(v) {
    const n = Number(String(v || '').trim());
    return Number.isFinite(n) ? n : null;
  }
}

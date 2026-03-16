/**
 * XLSX 文件解析器
 * 自动识别两种已知题库格式，AI 兜底未知格式
 */
class XlsxParser {
  /**
   * 解析文件 ArrayBuffer，返回标准化题目数组
   * @param {ArrayBuffer} buffer - 文件二进制数据
   * @param {Object} [options] - 选项
   * @param {Function} [options.onAIFallback] - AI 兜底回调，接收 (headers, sampleRows)，返回列映射
   * @returns {Promise<{questions: Array, formatType: string}>}
   */
  static async parse(buffer, options = {}) {
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    // 过滤空行
    const dataRows = rows.filter(row => row.some(cell => String(cell).trim()));
    if (dataRows.length < 2) {
      throw new Error('文件内容为空或行数不足');
    }

    const headers = dataRows[0].map(h => String(h).trim());
    const bodyRows = dataRows.slice(1);

    // 尝试格式1
    const format1 = XlsxParser._tryFormat1(headers, bodyRows);
    if (format1) return format1;

    // 尝试格式2
    const format2 = XlsxParser._tryFormat2(headers, bodyRows);
    if (format2) return format2;

    // AI 兜底
    if (options.onAIFallback) {
      const sampleRows = bodyRows.slice(0, 5);
      const mapping = await options.onAIFallback(headers, sampleRows);
      if (mapping) {
        return XlsxParser._parseWithMapping(headers, bodyRows, mapping);
      }
    }

    throw new Error('无法识别题库格式，且 AI 解析失败');
  }

  /**
   * 格式1: 序号 | 题型 | 题干 | 选项(A-xxx|B-xxx) | 答案
   * 特征: 有一列包含 | 或 - 分隔的选项文本
   */
  static _tryFormat1(headers, rows) {
    // 查找选项合并列（包含 A- 或 A. 开头的多选项文本）
    const sampleRow = rows[0];
    let optionColIdx = -1;
    let stemColIdx = -1;
    let typeColIdx = -1;
    let answerColIdx = -1;

    for (let i = 0; i < headers.length; i++) {
      const h = headers[i].toLowerCase();
      const val = String(sampleRow[i] || '');

      if (h.includes('题型') || h.includes('类型') || h === 'type') {
        typeColIdx = i;
      } else if (h.includes('题干') || h.includes('题目') || h.includes('题面')) {
        stemColIdx = i;
      } else if (h.includes('答案') || h.includes('正确') || h === 'answer') {
        answerColIdx = i;
      } else if (/[A-D][-.、]/.test(val) && (val.includes('|') || val.includes('\n') || /[A-D][-.、].*[A-D][-.、]/.test(val))) {
        optionColIdx = i;
      }
    }

    // 如果没通过表头找到，尝试通过内容推断
    if (stemColIdx === -1) {
      for (let i = 0; i < sampleRow.length; i++) {
        const val = String(sampleRow[i] || '');
        if (i !== optionColIdx && i !== typeColIdx && i !== answerColIdx && val.length > 10) {
          stemColIdx = i;
          break;
        }
      }
    }

    if (answerColIdx === -1) {
      // 答案列通常在最后，内容为短字母
      for (let i = headers.length - 1; i >= 0; i--) {
        const val = String(sampleRow[i] || '').trim();
        if (i !== optionColIdx && i !== stemColIdx && i !== typeColIdx && /^[A-Z,;、\s]+$/.test(val)) {
          answerColIdx = i;
          break;
        }
      }
    }

    if (stemColIdx === -1 || answerColIdx === -1) return null;

    const questions = [];
    for (const row of rows) {
      const stem = String(row[stemColIdx] || '').trim();
      if (!stem) continue;

      const rawType = typeColIdx >= 0 ? String(row[typeColIdx] || '').trim() : '';
      const type = XlsxParser._normalizeType(rawType);
      const rawAnswer = String(row[answerColIdx] || '').trim();
      const options = optionColIdx >= 0 ? XlsxParser._parseMergedOptions(String(row[optionColIdx] || '')) : [];

      questions.push(XlsxParser._buildQuestion(stem, type, options, rawAnswer));
    }

    return { questions, formatType: 'format1' };
  }

  /**
   * 格式2: 题型 | 题目标题 | 选项1(A) | 选项2(B) | ... | 正确答案
   * 特征: 多列分别是选项内容，最后一列是答案
   */
  static _tryFormat2(headers, rows) {
    // 寻找选项列模式（连续列，表头包含 A/B/C/D 或 选项1/选项2）
    let typeColIdx = -1;
    let stemColIdx = -1;
    let answerColIdx = -1;
    let optionStartIdx = -1;
    let optionEndIdx = -1;

    for (let i = 0; i < headers.length; i++) {
      const h = headers[i];
      if (/题型|类型/.test(h)) typeColIdx = i;
      else if (/题目|题干|题面|标题/.test(h)) stemColIdx = i;
      else if (/正确答案|答案|标准答案/.test(h)) answerColIdx = i;
      else if (/^[Aa]$|选项\s*1|选项\s*[Aa]|^A选项/.test(h)) optionStartIdx = i;
    }

    if (optionStartIdx === -1) {
      // 尝试检测：连续几列表头为 A/B/C/D
      for (let i = 0; i < headers.length - 1; i++) {
        if (/^[Aa]$/.test(headers[i]) && /^[Bb]$/.test(headers[i + 1])) {
          optionStartIdx = i;
          break;
        }
      }
    }

    if (stemColIdx === -1 || answerColIdx === -1 || optionStartIdx === -1) return null;

    // 确定选项列范围
    optionEndIdx = answerColIdx - 1;
    if (optionEndIdx < optionStartIdx) optionEndIdx = optionStartIdx + 3; // 默认4个选项

    const optionLabels = 'ABCDEFGHIJ';
    const questions = [];

    for (const row of rows) {
      const stem = String(row[stemColIdx] || '').trim();
      if (!stem) continue;

      const rawType = typeColIdx >= 0 ? String(row[typeColIdx] || '').trim() : '';
      const type = XlsxParser._normalizeType(rawType);
      const rawAnswer = String(row[answerColIdx] || '').trim();

      const options = [];
      for (let i = optionStartIdx; i <= optionEndIdx && i < row.length; i++) {
        const text = String(row[i] || '').trim();
        if (text) {
          options.push({
            label: optionLabels[i - optionStartIdx] || String.fromCharCode(65 + i - optionStartIdx),
            text
          });
        }
      }

      questions.push(XlsxParser._buildQuestion(stem, type, options, rawAnswer));
    }

    return { questions, formatType: 'format2' };
  }

  /**
   * 使用 AI 返回的列映射解析
   * @param {Array} headers
   * @param {Array} rows
   * @param {Object} mapping - { stem, type?, options?, optionStart?, optionEnd?, answer }
   */
  static _parseWithMapping(headers, rows, mapping) {
    const getIdx = (name) => {
      if (typeof name === 'number') return name;
      return headers.findIndex(h => h === name || h.includes(name));
    };

    const stemIdx = getIdx(mapping.stem);
    const typeIdx = mapping.type != null ? getIdx(mapping.type) : -1;
    const answerIdx = getIdx(mapping.answer);

    if (stemIdx === -1 || answerIdx === -1) {
      throw new Error('AI 映射的列名无法在表头中找到');
    }

    const optionLabels = 'ABCDEFGHIJ';
    const questions = [];

    for (const row of rows) {
      const stem = String(row[stemIdx] || '').trim();
      if (!stem) continue;

      const rawType = typeIdx >= 0 ? String(row[typeIdx] || '').trim() : '';
      const type = XlsxParser._normalizeType(rawType);
      const rawAnswer = String(row[answerIdx] || '').trim();

      let options = [];
      if (mapping.options != null) {
        // 选项合并在一列
        const optIdx = getIdx(mapping.options);
        if (optIdx >= 0) options = XlsxParser._parseMergedOptions(String(row[optIdx] || ''));
      } else if (mapping.optionStart != null) {
        // 选项分列
        const startIdx = getIdx(mapping.optionStart);
        const endIdx = mapping.optionEnd != null ? getIdx(mapping.optionEnd) : answerIdx - 1;
        for (let i = startIdx; i <= endIdx && i < row.length; i++) {
          const text = String(row[i] || '').trim();
          if (text) {
            options.push({ label: optionLabels[i - startIdx], text });
          }
        }
      }

      questions.push(XlsxParser._buildQuestion(stem, type, options, rawAnswer));
    }

    return { questions, formatType: 'ai_mapped' };
  }

  /** 解析合并选项文本 "A-xxx|B-xxx" 或 "A.xxx\nB.xxx" */
  static _parseMergedOptions(text) {
    if (!text.trim()) return [];
    // 尝试按 | 或换行分割
    let parts = text.split(/[|\n]/);
    if (parts.length <= 1) {
      // 尝试按 "B-" "C-" 等位置分割
      parts = text.split(/(?=[A-Z][-.、])/).filter(Boolean);
    }

    return parts.map(part => {
      const trimmed = part.trim();
      const match = trimmed.match(/^([A-Z])[-.、:：\s]\s*(.*)/);
      if (match) return { label: match[1], text: match[2].trim() };
      return null;
    }).filter(Boolean);
  }

  /** 标准化题型 */
  static _normalizeType(raw) {
    if (/单[选择]|single/i.test(raw)) return 'single';
    if (/多[选择]|multiple/i.test(raw)) return 'multiple';
    if (/判断|是非|对错|true.*false|bool/i.test(raw)) return 'judge';
    if (/填空|blank|fill/i.test(raw)) return 'fill';
    return 'single'; // 默认单选
  }

  /** 构建标准化题目对象 */
  static _buildQuestion(stem, type, options, rawAnswer) {
    const id = XlsxParser._hashStem(stem);
    const stemNormalized = XlsxParser.normalize(stem);

    // 解析答案为数组
    const answerArray = XlsxParser._parseAnswer(rawAnswer, type);

    return { id, type, stem, stemNormalized, options, answer: rawAnswer, answerArray };
  }

  /** 解析答案字符串为数组 */
  static _parseAnswer(raw, type) {
    if (!raw) return [];
    const cleaned = raw.replace(/[,;、，；\s]/g, '');
    if (type === 'judge') {
      if (/^[对是√正确truecorrect✓]+$/i.test(cleaned)) return ['对'];
      if (/^[错否×不正确falsewrong✗]+$/i.test(cleaned)) return ['错'];
      return [cleaned];
    }
    // 选择题：拆分为单个字母
    const letters = cleaned.match(/[A-Z]/gi);
    if (letters) return letters.map(l => l.toUpperCase());
    return [raw.trim()];
  }

  /** 归一化题干：去题号、去标点、去空格、统一小写 */
  static normalize(text) {
    return text
      .replace(/^[\d\s.、）)]+/, '')       // 去题号
      .replace(/[（()）【】\[\]《》<>""''「」『』]/g, '') // 去括号引号
      .replace(/[,;.!?，；。！？：:、\s\u3000\t\r\n]+/g, '') // 去标点空格
      .toLowerCase();
  }

  /** 简单字符串哈希 */
  static _hashStem(stem) {
    let hash = 0;
    const s = XlsxParser.normalize(stem);
    for (let i = 0; i < s.length; i++) {
      hash = ((hash << 5) - hash + s.charCodeAt(i)) | 0;
    }
    return 'q_' + Math.abs(hash).toString(36);
  }
}

// 暴露为全局变量（content script 环境）
window.XlsxParser = XlsxParser;

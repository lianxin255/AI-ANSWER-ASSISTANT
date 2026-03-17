// Content Script - 题目识别和自动答题

(function () {
  "use strict";

  // State
  let questions = [];
  let answeredCount = 0;
  let isRunning = false;
  let config = null;
  let answerStrategy = "local_first";

  const ANSWER_STRATEGY = {
    LOCAL_FIRST: "local_first",
    AI_ONLY: "ai_only",
    LOCAL_ONLY: "local_only",
  };

  // Question selectors for common exam platforms
  const QUESTION_SELECTORS = [
    // 通用选择器
    ".question",
    ".question-item",
    ".exam-question",
    ".test-question",
    ".quiz-question",
    '[class*="question"]',
    '[class*="Question"]',
    // 题目容器
    ".problem",
    ".problem-item",
    ".exercise",
    ".exercise-item",
    // 表单题目
    "form .item",
    "form .form-item",
    // 列表题目
    ".question-list > li",
    ".question-list > div",
    "ol.questions > li",
    "ul.questions > li",
  ];

  // Option selectors
  const OPTION_SELECTORS = [
    'input[type="radio"]',
    'input[type="checkbox"]',
    ".option",
    ".choice",
    ".answer-option",
    '[class*="option"]',
    '[class*="choice"]',
    "label",
  ];

  // Fill-in-the-blank selectors
  const FILL_SELECTORS = [
    'input[type="text"]',
    "input:not([type])",
    "textarea",
    ".blank",
    ".fill-blank",
    '[class*="blank"]',
    '[contenteditable="true"]',
  ];

  // AI分析得到的选择器缓存
  let aiDetectedSelectors = null;

  // Message listener
  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    switch (message.action) {
      case "scan":
        config = message.config;
        handleScan(sendResponse);
        return true; // 保持消息通道开放用于异步响应
      case "start":
        config = message.config;
        startAnswering();
        sendResponse({ success: true });
        break;
      case "stop":
        stopAnswering();
        sendResponse({ success: true });
        break;
      case "getStatus":
        sendResponse({
          questionCount: questions.length,
          answeredCount,
          isRunning,
        });
        break;
    }
    return true;
  });

  // 处理扫描请求
  async function handleScan(sendResponse) {
    if (!config) {
      sendResponse({ success: false, count: 0, message: "请先配置API" });
      return;
    }

    // 步骤1: 尝试匹配站点模板
    sendLog("info", "正在匹配站点模板...");
    const template = window.siteMatcher.matchTemplate(window.location.href);

    if (template) {
      sendLog("info", `已匹配到站点模板: ${template.siteName}`);

      try {
        // 使用模板扫描
        const scanner = new window.EnhancedScanner();
        const result = scanner.scanWithTemplate(template);

        if (result.success && result.count > 0) {
          // 模板扫描成功
          questions = result.questions;
          answeredCount = 0;
          updateStats();

          sendLog("success", `使用模板扫描成功，发现 ${result.count} 道题目`);

          // 更新模板统计
          await window.templateManager.updateStats(template.siteId, "success");

          sendResponse({ success: true, count: result.count, message: "" });
          return;
        } else {
          // 模板扫描失败，回退到AI分析
          sendLog("warning", `模板扫描失败，回退到AI分析...`);
          await window.templateManager.updateStats(template.siteId, "fail");
        }
      } catch (error) {
        sendLog("error", `模板扫描出错: ${error.message}，回退到AI分析`);
        await window.templateManager.updateStats(template.siteId, "fail");
      }
    } else {
      sendLog("info", "未找到匹配的站点模板，使用AI分析...");
    }

    // 步骤2: 使用AI分析（无模板或模板失败时）
    sendLog("info", "正在使用AI分析页面结构，请耐心等待...");

    try {
      const aiResult = await analyzePageWithAI();
      if (aiResult && aiResult.success) {
        aiDetectedSelectors = aiResult.selectors;
        const count = scanWithAISelectors(aiResult);
        if (count > 0) {
          sendLog("success", `AI分析成功，发现 ${count} 道题目`);
          sendResponse({ success: true, count, message: "" });
        } else {
          sendLog("warning", "AI分析完成，但未能定位到题目元素");
          sendResponse({
            success: false,
            count: 0,
            message: "未能定位到题目元素",
          });
        }
      } else {
        sendLog("warning", "AI分析未发现题目");
        sendResponse({ success: false, count: 0, message: "未发现题目" });
      }
    } catch (error) {
      sendLog("error", `AI分析失败: ${error.message}`);
      sendResponse({ success: false, count: 0, message: error.message });
    }
  }

  // Scan for questions on the page
  function scanQuestions() {
    questions = [];
    answeredCount = 0;

    // Try each selector
    for (const selector of QUESTION_SELECTORS) {
      try {
        const elements = document.querySelectorAll(selector);
        if (elements.length > 0) {
          elements.forEach((el, index) => {
            const question = parseQuestion(el, index);
            if (question) {
              questions.push(question);
            }
          });
          if (questions.length > 0) break;
        }
      } catch (e) {
        console.log("Selector error:", selector, e);
      }
    }

    // If no questions found with selectors, try heuristic detection
    if (questions.length === 0) {
      questions = detectQuestionsHeuristically();
    }

    // Remove duplicates
    questions = removeDuplicates(questions);

    sendLog("info", `扫描完成，发现 ${questions.length} 道题目`);
    updateStats();

    return questions.length;
  }

  // Parse a question element
  function parseQuestion(element, index) {
    const question = {
      index,
      element,
      type: null,
      text: "",
      options: [],
      inputs: [],
      answered: false,
    };

    // Get question text
    const textElements = element.querySelectorAll(
      "p, span, div, h1, h2, h3, h4, h5, h6"
    );
    let questionText = "";

    // Try to find the main question text
    const titleEl = element.querySelector(
      '.title, .question-title, .question-text, .stem, [class*="title"], [class*="stem"]'
    );
    if (titleEl) {
      questionText = titleEl.textContent.trim();
    } else {
      // Get first meaningful text
      for (const el of textElements) {
        const text = el.textContent.trim();
        if (text.length > 10 && !text.match(/^[A-D][\.\、\s]/)) {
          questionText = text;
          break;
        }
      }
    }

    if (!questionText) {
      questionText = element.textContent.trim().substring(0, 500);
    }

    question.text = cleanText(questionText);

    // Detect question type and get options/inputs
    const radios = element.querySelectorAll('input[type="radio"]');
    const checkboxes = element.querySelectorAll('input[type="checkbox"]');
    const textInputs = element.querySelectorAll(
      'input[type="text"], input:not([type]), textarea'
    );

    if (radios.length > 0) {
      question.type = "single";
      question.options = parseOptions(element, radios);
    } else if (checkboxes.length > 0) {
      question.type = "multiple";
      question.options = parseOptions(element, checkboxes);
    } else if (textInputs.length > 0) {
      question.type = "fill";
      question.inputs = Array.from(textInputs);
    } else {
      // Try to detect from text
      if (
        question.text.includes("多选") ||
        question.text.includes("多项选择")
      ) {
        question.type = "multiple";
      } else if (
        question.text.includes("单选") ||
        question.text.includes("单项选择")
      ) {
        question.type = "single";
      } else if (
        question.text.includes("填空") ||
        question.text.includes("____") ||
        question.text.includes("___")
      ) {
        question.type = "fill";
      }

      // Try to find clickable options
      const optionEls = element.querySelectorAll(
        '.option, .choice, [class*="option"], [class*="choice"], li'
      );
      if (optionEls.length >= 2 && optionEls.length <= 8) {
        question.type = question.type || "single";
        question.options = Array.from(optionEls).map((el, i) => ({
          element: el,
          label: String.fromCharCode(65 + i),
          text: el.textContent.trim(),
        }));
      }
    }

    // Skip if no valid type detected
    if (!question.type) {
      return null;
    }

    return question;
  }

  // Parse options from input elements
  function parseOptions(container, inputs) {
    const options = [];

    inputs.forEach((input, index) => {
      const label =
        input.closest("label") ||
        container.querySelector(`label[for="${input.id}"]`) ||
        input.parentElement;

      let optionText = "";
      let optionLabel = String.fromCharCode(65 + index);

      if (label) {
        optionText = label.textContent.trim();
        // Extract label letter if present
        const match = optionText.match(/^([A-Z])[\.\、\s]/);
        if (match) {
          optionLabel = match[1];
          optionText = optionText.substring(match[0].length).trim();
        }
      }

      options.push({
        element: input,
        label: optionLabel,
        text: optionText,
      });
    });

    return options;
  }

  // Heuristic question detection
  function detectQuestionsHeuristically() {
    const detected = [];
    const allElements = document.body.querySelectorAll("*");

    // Look for numbered items that might be questions
    const numberPattern = /^[\d一二三四五六七八九十]+[\.\、\s]/;

    allElements.forEach((el, index) => {
      const text = el.textContent.trim();
      if (text.length > 20 && text.length < 2000 && numberPattern.test(text)) {
        // Check if it has options or inputs
        const hasOptions =
          el.querySelectorAll('input[type="radio"], input[type="checkbox"]')
            .length > 0;
        const hasInputs =
          el.querySelectorAll('input[type="text"], textarea').length > 0;

        if (hasOptions || hasInputs) {
          const question = parseQuestion(el, detected.length);
          if (question) {
            detected.push(question);
          }
        }
      }
    });

    return detected;
  }

  // 统一的DOM精简逻辑，去掉无关属性
  function simplifyElementAttributes(el) {
    const keepAttrs = [
      "class",
      "id",
      "type",
      "name",
      "value",
      "placeholder",
      "for",
      "data-index",
      "data-id",
    ];
    const attrs = Array.from(el.attributes || []);
    attrs.forEach((attr) => {
      if (!keepAttrs.includes(attr.name) && !attr.name.startsWith("data-")) {
        el.removeAttribute(attr.name);
      }
    });
    Array.from(el.children).forEach((child) =>
      simplifyElementAttributes(child)
    );
  }

  // 获取简化的页面HTML用于AI分析（作为候选块不足时的兜底）
  function getSimplifiedHTML() {
    const clone = document.body.cloneNode(true);
    const removeSelectors = [
      "script",
      "style",
      "noscript",
      "iframe",
      "svg",
      "img",
      "video",
      "audio",
      "canvas",
    ];
    removeSelectors.forEach((sel) => {
      clone.querySelectorAll(sel).forEach((el) => el.remove());
    });

    simplifyElementAttributes(clone);

    let html = clone.innerHTML;
    html = html.replace(/\s+/g, " ").replace(/>\s+</g, "><");

    if (html.length > 30000) {
      const mainSelectors = [
        "main",
        "article",
        ".content",
        ".main",
        "#content",
        "#main",
        ".container",
        ".wrapper",
      ];
      for (const sel of mainSelectors) {
        const main = clone.querySelector(sel);
        if (main && main.innerHTML.length > 500) {
          html = main.innerHTML.replace(/\s+/g, " ").replace(/>\s+</g, "><");
          break;
        }
      }
    }

    if (html.length > 30000) {
      html = html.substring(0, 30000) + "... [内容已截断]";
    }

    return html;
  }

  // 构建单个题目块的精简HTML字符串
  function buildSimplifiedBlockHTML(element, maxLength = 1500) {
    const clone = element.cloneNode(true);
    simplifyElementAttributes(clone);
    let html = clone.outerHTML || "";
    html = html.replace(/\s+/g, " ").replace(/>\s+</g, "><");
    if (html.length > maxLength) {
      html = html.substring(0, maxLength) + "... [片段截断]";
    }
    return html;
  }

  // 根据页面内容尝试提取候选题目块，显著减少发送给AI的数据量
  function getCandidateQuestionBlocks(maxBlocks = 40) {
    const candidateSet = new Set();

    function addCandidate(el) {
      if (!el || candidateSet.has(el)) return;
      candidateSet.add(el);
    }

    const optionInputs = document.querySelectorAll(
      'input[type="radio"], input[type="checkbox"]'
    );
    optionInputs.forEach((input) => {
      const container = findQuestionContainer(input.closest("label") || input);
      if (container) addCandidate(container);
    });

    const fillInputs = document.querySelectorAll(
      'input[type="text"], input[type="number"], textarea'
    );
    fillInputs.forEach((input) => {
      const container = findQuestionContainer(input.closest("label") || input);
      if (container) addCandidate(container);
    });

    if (candidateSet.size < 5) {
      const extraSelectors = [
        ".question",
        ".exam-question",
        ".topic",
        ".subject",
        ".problem",
      ];
      extraSelectors.forEach((sel) => {
        document.querySelectorAll(sel).forEach((el) => addCandidate(el));
      });
    }

    const candidates = Array.from(candidateSet).slice(0, maxBlocks);
    return candidates
      .map((el) => {
        const text = cleanText(el.textContent || "").substring(0, 400);
        return {
          text,
          html: buildSimplifiedBlockHTML(el),
        };
      })
      .filter((block) => block.text);
  }

  // 构建AI分析所需的内容，如果候选块为空则回退到整页HTML
  function buildAIAnalysisPayload() {
    const blocks = getCandidateQuestionBlocks();
    if (blocks.length > 0) {
      const formatted = blocks
        .map((block, index) => {
          return `【题目块${index + 1}】\n文本：${block.text}\nHTML：${
            block.html
          }`;
        })
        .join("\n\n");

      return {
        payload: `以下是筛选后的疑似题目区域（共${blocks.length}块）：\n${formatted}`,
        source: "candidateBlocks",
      };
    }

    return {
      payload: getSimplifiedHTML(),
      source: "fullHTML",
    };
  }

  // 使用AI分析页面结构（只识别题目，不返回答案）
  async function analyzePageWithAI() {
    const { payload, source } = buildAIAnalysisPayload();
    const contentIntro =
      source === "candidateBlocks"
        ? "本次提供的是经过前端筛选的疑似题目块，请基于这些块识别题目结构。"
        : "未找到足够的候选题目块，以下为整页精简HTML。";

    const prompt = `分析以下HTML页面，识别所有题目的结构。

【重要】只需要识别题目结构，不需要给出答案！

请返回JSON格式：
{
  "success": true,
  "questions": [
    {
      "index": 0,
      "type": "single",
      "text": "题目文本内容（完整的题干）",
      "options": [
        {
          "label": "A",
          "text": "选项内容",
          "selector": "精确的CSS选择器"
        }
      ]
    },
    {
      "index": 1,
      "type": "multiple",
      "text": "多选题文本",
      "options": [
        {
          "label": "A",
          "text": "选项内容",
          "selector": "CSS选择器"
        }
      ]
    },
    {
      "index": 2,
      "type": "fill",
      "text": "填空题文本",
      "inputs": [
        { "selector": "第1个输入框的CSS选择器" },
        { "selector": "第2个输入框的CSS选择器" }
      ]
    }
  ]
}

重要说明：
1. type: "single"单选题, "multiple"多选题, "fill"填空题
2. text: 必须包含完整的题干内容，后续需要用这个文本去获取答案
3. selector: 必须是可以直接用document.querySelector()定位到的精确CSS选择器
   - 优先使用id选择器: #elementId
   - 或使用class+nth-child: .option-item:nth-child(2)
   - 或使用属性选择器: input[value="B"], input[name="q1"][value="2"]
   - 对于radio/checkbox，选择器应指向input元素本身
   - 对于可点击的div/label，选择器应指向该可点击元素
4. 仔细分析HTML结构，确保选择器准确无误
5. 【不要返回答案】这一步只需要识别题目结构

${contentIntro}

HTML内容：
${payload}`;

    return new Promise((resolve, reject) => {
      chrome.runtime.sendMessage(
        {
          action: "analyzeHTML",
          config,
          prompt,
        },
        (response) => {
          if (chrome.runtime.lastError) {
            reject(new Error(chrome.runtime.lastError.message));
            return;
          }

          if (!response.success) {
            reject(new Error(response.error));
            return;
          }

          try {
            const result = parseAIResponse(response.data);
            resolve(result);
          } catch (e) {
            reject(new Error("解析AI响应失败: " + e.message));
          }
        }
      );
    });
  }

  // 使用AI分析结果创建题目列表（不含答案，答案在答题阶段逐题获取）
  function scanWithAISelectors(aiResult) {
    questions = [];
    answeredCount = 0;

    if (!aiResult.questions || aiResult.questions.length === 0) {
      return 0;
    }

    // 使用AI返回的题目结构数据（不含答案）
    aiResult.questions.forEach((q, index) => {
      const question = {
        index,
        type: q.type || "single",
        text: q.text || "",
        answer: null, // 答案在答题阶段获取
        explanation: null,
        options: [],
        inputs: [],
        answered: false,
      };

      // 处理选项（单选/多选题）
      if (q.options && q.options.length > 0) {
        question.options = q.options.map((opt) => ({
          label: opt.label,
          text: opt.text,
          selector: opt.selector,
          element: opt.selector ? safeQuerySelector(opt.selector) : null,
        }));
      }

      // 处理填空题输入框
      if (q.type === "fill" && q.inputs && q.inputs.length > 0) {
        question.inputs = q.inputs.map((inp) => ({
          selector: inp.selector,
          element: inp.selector ? safeQuerySelector(inp.selector) : null,
        }));
      }

      questions.push(question);
      console.log(
        `[AI答题助手] 题目${index + 1}:`,
        question.text.substring(0, 50)
      );
    });

    updateStats();
    return questions.length;
  }

  // 安全的querySelector，捕获无效选择器错误
  function safeQuerySelector(selector) {
    if (!selector) return null;
    try {
      return document.querySelector(selector);
    } catch (e) {
      console.warn("[AI答题助手] 无效的选择器:", selector, e);
      return null;
    }
  }

  // Remove duplicate questions
  function removeDuplicates(questions) {
    const seen = new Set();
    return questions.filter((q) => {
      const key = q.text.substring(0, 100);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  // Clean text
  function cleanText(text) {
    return text
      .replace(/\s+/g, " ")
      .replace(/[\r\n]+/g, " ")
      .trim()
      .substring(0, 1000);
  }

  // Start answering questions
  async function startAnswering() {
    if (isRunning) return;

    isRunning = true;

    if (questions.length === 0) {
      sendLog("warning", "请先扫描题目");
      isRunning = false;
      sendComplete();
      return;
    }

    answerStrategy = await loadAnswerStrategy();
    sendLog("info", `开始答题，共 ${questions.length} 道题目`);
    sendLog("info", `答题策略：${getStrategyLabel(answerStrategy)}`);

    // 加载关联题库（策略允许时）
    if (answerStrategy !== ANSWER_STRATEGY.AI_ONLY) {
      try {
        const url = window.location.href;
        const binding = await QuestionBankManager.findBindingByUrl(url);
        if (binding && binding.activeBankIds && binding.activeBankIds.length > 0) {
          await window.questionBankMatcher.loadBanks(binding.activeBankIds);
          sendLog("info", `已加载本地题库（${window.questionBankMatcher.questionCount} 题），优先本地匹配`);
        }
      } catch (e) {
        console.warn("[AI答题助手] 题库加载失败，将全部使用 AI:", e);
      }
    } else if (window.questionBankMatcher) {
      window.questionBankMatcher.unload();
    }

    for (let i = 0; i < questions.length; i++) {
      if (!isRunning) {
        sendLog("warning", "答题已停止");
        break;
      }

      const question = questions[i];

      if (question.answered) {
        continue;
      }

      // 找到第一个有效的选项元素用于滚动定位和高亮
      const firstElement = findFirstValidElement(question);

      if (firstElement) {
        // 滚动到题目位置
        scrollToElement(firstElement);
        await sleep(300);

        // 高亮当前题目区域
        const questionContainer = findQuestionContainer(firstElement);
        if (questionContainer) {
          highlightElement(questionContainer);
        }
      }

      sendLog(
        "info",
        `正在处理第 ${i + 1}/${questions.length} 题: ${question.text.substring(
          0,
          30
        )}...`
      );

      try {
        // 先尝试本地题库匹配
        let answer = null;
        let isLocalMatch = false;
        if (answerStrategy !== ANSWER_STRATEGY.AI_ONLY && window.questionBankMatcher && window.questionBankMatcher.isLoaded()) {
          const match = window.questionBankMatcher.match(question);
          if (match) {
            // 转换题库答案为可应用的格式
            answer = convertBankAnswer(match.question, question);
            isLocalMatch = true;
            sendLog("success", `第 ${i + 1} 题本地匹配成功（级别${match.level}，相似度${(match.score * 100).toFixed(0)}%）`);
          }
        }

        // 本地未匹配，调用 AI
        if (!answer) {
          if (answerStrategy === ANSWER_STRATEGY.LOCAL_ONLY) {
            sendLog("warning", `第 ${i + 1} 题本地未匹配，按策略跳过`);
            continue;
          }
          sendLog("info", `正在获取第 ${i + 1} 题的答案...`);
          answer = await getAIAnswerForQuestion(question);
        }

        if (answer && answer.answer) {
          question.answer = answer.answer;
          question.explanation = answer.explanation;
          await applyAnswerDirectly(question);
          question.answered = true;
          answeredCount++;
          // 更新统计
          updateStats();

          // 发送统计
          chrome.runtime.sendMessage({
            action: "trackStats",
            event: "question_answered",
          });
          sendLog(
            "success",
            `第 ${i + 1} 题已完成${isLocalMatch ? '(本地)' : '(AI)'}，答案: ${JSON.stringify(question.answer)}`
          );
        } else {
          sendLog("warning", `第 ${i + 1} 题未能获取答案`);
        }
      } catch (error) {
        sendLog("error", `第 ${i + 1} 题处理失败: ${error.message}`);
        console.error("[AI答题助手] 答题错误:", error);
      }

      // 移除高亮
      if (firstElement) {
        const questionContainer = findQuestionContainer(firstElement);
        if (questionContainer) {
          removeHighlight(questionContainer);
          // 添加已完成标记
          markAsCompleted(questionContainer);
        }
      }

      // Wait before next question
      await sleep(500);
    }

    // 批量更新题库匹配统计
    if (window.questionBankMatcher && window.questionBankMatcher.isLoaded()) {
      try {
        const stats = window.questionBankMatcher.getMatchStats();
        for (const [bankId, count] of Object.entries(stats)) {
          if (count > 0) {
            await QuestionBankManager.updateBankStats(bankId, count);
          }
        }
      } catch (_e) { /* 统计更新失败不影响主流程 */ }
    }

    isRunning = false;
    sendComplete();
  }

  // 找到题目中第一个有效的元素用于定位
  function findFirstValidElement(question) {
    if (question.options && question.options.length > 0) {
      for (const opt of question.options) {
        if (opt.element) return opt.element;
        // 尝试重新查询
        if (opt.selector) {
          const el = safeQuerySelector(opt.selector);
          if (el) return el;
        }
      }
    }
    if (question.inputs && question.inputs.length > 0) {
      for (const inp of question.inputs) {
        if (inp.element) return inp.element;
        if (inp.selector) {
          const el = safeQuerySelector(inp.selector);
          if (el) return el;
        }
      }
    }
    return null;
  }

  async function loadAnswerStrategy() {
    try {
      const result = await chrome.storage.sync.get(["answerStrategy"]);
      const value = result.answerStrategy;
      if (Object.values(ANSWER_STRATEGY).includes(value)) {
        return value;
      }
    } catch (_e) {
      // ignore
    }
    return ANSWER_STRATEGY.LOCAL_FIRST;
  }

  function getStrategyLabel(strategy) {
    switch (strategy) {
      case ANSWER_STRATEGY.AI_ONLY:
        return "仅 AI";
      case ANSWER_STRATEGY.LOCAL_ONLY:
        return "仅题库";
      default:
        return "优先题库，未命中用 AI";
    }
  }

  // 根据选项元素向上查找题目容器（精确定位到单道题）
  function findQuestionContainer(element) {
    if (!element) return null;

    // 优先使用腾讯问卷的精确容器选择器
    const tencentContainer = element.closest(
      "section.question[data-question-id]"
    );
    if (tencentContainer) {
      return tencentContainer;
    }

    // 先找到这个选项所属的所有同级选项（同一道题的选项）
    const elementInput =
      element.tagName === "INPUT" ? element : element.querySelector("input");
    const inputName = elementInput?.name;

    let current = element.parentElement;
    let bestContainer = null;
    let depth = 0;
    const maxDepth = 4; // 限制层数，只找最近的容器

    while (current && current !== document.body && depth < maxDepth) {
      // 检查当前容器内有多少组选项（通过不同的 name 判断）
      const allInputs = current.querySelectorAll(
        'input[type="radio"], input[type="checkbox"]'
      );
      const names = new Set();
      allInputs.forEach((inp) => {
        if (inp.name) names.add(inp.name);
      });

      // 如果这个容器只包含一道题的选项（1个name），就是我们要的
      if (names.size === 1 && allInputs.length >= 2) {
        bestContainer = current;
        // 继续向上找一层，看看父元素是否也只包含这一道题
        // 但不要找太多层
      } else if (names.size > 1) {
        // 包含多道题了，停止，使用上一个找到的容器
        break;
      }

      // 如果没有 input，检查是否有可点击的选项元素
      if (allInputs.length === 0) {
        const options = current.querySelectorAll(
          '.option, [class*="option"], [class*="choice"]'
        );
        if (options.length >= 2 && options.length <= 6) {
          bestContainer = current;
        }
      }

      current = current.parentElement;
      depth++;
    }

    // 如果没找到，返回选项的直接父元素的父元素
    return (
      bestContainer ||
      element.parentElement?.parentElement ||
      element.parentElement ||
      element
    );
  }

  // 标记题目为已完成
  function markAsCompleted(element) {
    if (!element) return;

    // 添加完成样式
    // 检查是否已有标记
    if (element.dataset.aiAnswerCompleted === "true") return;

    // 创建完成标记
    removeHighlight(element);
    element.dataset.aiAnswerCompleted = "true";
    element.style.outline = "2px solid #22c55e";
    element.style.outlineOffset = "2px";
  }

  // 直接应用答案（使用AI返回的选择器）
  async function applyAnswerDirectly(question) {
    switch (question.type) {
      case "single":
      case "judge":
        await applySingleAnswerDirectly(question);
        break;
      case "multiple":
        await applyMultipleAnswerDirectly(question);
        break;
      case "fill":
        await applyFillAnswerDirectly(question);
        break;
    }
  }

  // 单选题 - 直接点击对应选项
  async function applySingleAnswerDirectly(question) {
    const answerLetter = String(question.answer).toUpperCase();

    for (const option of question.options) {
      if (option.label.toUpperCase() === answerLetter) {
        // 优先使用已解析的element
        let element = option.element;

        // 如果element无效，尝试重新查询
        if (!element && option.selector) {
          element = safeQuerySelector(option.selector);
        }

        if (element) {
          await clickElement(element);
          console.log(`[AI答题助手] 单选已点击: ${option.label}`);
        } else {
          console.warn(
            `[AI答题助手] 找不到选项元素: ${option.label}, selector: ${option.selector}`
          );
        }
        break;
      }
    }
  }

  // 多选题 - 点击所有正确选项
  async function applyMultipleAnswerDirectly(question) {
    let answers = question.answer;

    // 确保answers是数组
    if (typeof answers === "string") {
      answers = answers
        .split("")
        .filter((c) => /[A-Z]/i.test(c))
        .map((c) => c.toUpperCase());
    } else if (Array.isArray(answers)) {
      answers = answers.map((a) => String(a).toUpperCase());
    }

    console.log(`[AI答题助手] 多选答案:`, answers, `| 题目: ${question.text.substring(0, 30)}`);
    console.log(`[AI答题助手] 可用选项:`, question.options.map(o => `${o.label}=${o.text}`));

    for (const option of question.options) {
      if (answers.includes(option.label.toUpperCase())) {
        let element = option.element;

        if (!element && option.selector) {
          element = safeQuerySelector(option.selector);
        }

        if (element) {
          await clickElement(element);
          console.log(`[AI答题助手] 多选已点击: ${option.label}`);
          await sleep(200);
        } else {
          console.warn(
            `[AI答题助手] 找不到选项元素: ${option.label}, selector: ${option.selector}`
          );
        }
      }
    }
  }

  // 填空题 - 填写答案
  async function applyFillAnswerDirectly(question) {
    let answers = question.answer;

    // 确保answers是数组
    if (!Array.isArray(answers)) {
      answers = [answers];
    }

    // 获取第一个输入框
    let inputElement = null;

    for (const inputInfo of question.inputs) {
      let element = inputInfo.element;
      if (!element && inputInfo.selector) {
        element = safeQuerySelector(inputInfo.selector);
      }
      if (element) {
        inputElement = element;
        break;
      }
    }

    if (!inputElement) {
      console.warn(`[AI答题助手] 找不到填空题输入框`);
      return;
    }

    // 将所有答案合并为一个字符串填入同一个输入框
    // 如果只有一个答案，直接使用；多个答案用中文分号 "；" 分隔
    const combinedAnswer =
      answers.length === 1
        ? String(answers[0])
        : answers.map((a) => String(a)).join("；");

    await fillInput(inputElement, combinedAnswer);
    console.log(`[AI答题助手] 填空已填写: ${combinedAnswer}`);
  }

  // 填写输入框
  async function fillInput(element, value) {
    element.focus();
    await sleep(50);

    // 清空现有值
    element.value = "";

    // 设置新值
    element.value = value;

    // 触发各种事件确保框架能检测到变化
    element.dispatchEvent(new Event("input", { bubbles: true }));
    element.dispatchEvent(new Event("change", { bubbles: true }));
    element.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true }));
    element.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true }));

    element.blur();
  }

  // Stop answering
  function stopAnswering() {
    isRunning = false;
  }

  // 逐题获取AI答案（只发送单道题目，不发送整页HTML）
  async function getAIAnswerForQuestion(question) {
    let prompt = `请回答以下${getTypeLabel(question.type)}：\n\n`;
    prompt += `题目：${question.text}\n\n`;

    if (question.options && question.options.length > 0) {
      prompt += "选项：\n";
      question.options.forEach((opt) => {
        prompt += `${opt.label}. ${opt.text}\n`;
      });
      prompt += "\n";
    }

    if (
      question.type === "fill" &&
      question.inputs &&
      question.inputs.length > 1
    ) {
      prompt += `（共有 ${question.inputs.length} 个空需要填写）\n\n`;
    }

    prompt += `请严格按照JSON格式返回答案：
{
  "type": "${question.type}",
  "answer": ${
    question.type === "single"
      ? '"选项字母如B"'
      : question.type === "multiple"
      ? '["A", "C"]'
      : '["答案1", "答案2"]'
  },
  "explanation": "简短解释"
}

注意：
- 单选题answer为单个字母，如 "B"
- 多选题answer为字母数组，如 ["A", "C"]
- 填空题answer为答案数组，如 ["答案1"] 或 ["答案1", "答案2"]
- 只返回JSON，不要其他内容`;

    return new Promise((resolve, reject) => {
      chrome.runtime.sendMessage(
        {
          action: "callAI",
          config,
          prompt,
        },
        (response) => {
          if (chrome.runtime.lastError) {
            reject(new Error(chrome.runtime.lastError.message));
            return;
          }

          if (!response.success) {
            reject(new Error(response.error));
            return;
          }

          try {
            const answer = parseAIResponse(response.data);
            resolve(answer);
          } catch (e) {
            reject(new Error("解析AI响应失败: " + e.message));
          }
        }
      );
    });
  }

  // Get AI answer for a question (legacy, kept for compatibility)
  async function getAIAnswer(question) {
    return getAIAnswerForQuestion(question);
  }

  // Build prompt for AI
  function buildPrompt(question) {
    let prompt = `题目类型: ${getTypeLabel(question.type)}\n\n`;
    prompt += `题目: ${question.text}\n\n`;

    if (question.options.length > 0) {
      prompt += "选项:\n";
      question.options.forEach((opt) => {
        prompt += `${opt.label}. ${opt.text}\n`;
      });
    }

    if (question.type === "fill") {
      prompt += `\n这是一道填空题，请给出填空的答案。`;
      if (question.inputs.length > 1) {
        prompt += `共有 ${question.inputs.length} 个空需要填写。`;
      }
    }

    return prompt;
  }

  // Get type label
  /**
   * 将题库匹配结果转换为可应用的答案格式
   * 处理: 判断题→单选转换、选项文本回退匹配、填空题直接传递
   */
  function convertBankAnswer(bankQuestion, scannedQuestion) {
    const bankType = bankQuestion.type;
    const scanType = scannedQuestion.type;
    const bankAnswerArray = bankQuestion.answerArray || [];
    const rawAnswer = bankQuestion.answer || '';

    // 1. 填空题：直接使用答案文本
    if (scanType === 'fill') {
      return {
        success: true,
        type: 'fill',
        answer: bankAnswerArray.length > 0 ? bankAnswerArray : [rawAnswer],
        explanation: '[本地题库]'
      };
    }

    // 2. 判断题→单选转换：题库答案是"对/错"，需要匹配页面选项
    if (bankType === 'judge' && scanType === 'single') {
      const isCorrect = /^[对是√正确true✓]+$/i.test(rawAnswer.replace(/[,;、，；\s]/g, ''));
      const isWrong = /^[错否×不正确false✗]+$/i.test(rawAnswer.replace(/[,;、，；\s]/g, ''));

      if (isCorrect || isWrong) {
        // 尝试通过选项文本匹配
        for (const opt of (scannedQuestion.options || [])) {
          const optText = (opt.text || '').trim();
          if (isCorrect && /^[对是正确√✓]/.test(optText)) {
            return { success: true, type: 'single', answer: opt.label, explanation: '[本地题库-判断]' };
          }
          if (isWrong && /^[错否不正确×✗]/.test(optText)) {
            return { success: true, type: 'single', answer: opt.label, explanation: '[本地题库-判断]' };
          }
        }
        // 常见约定：A=对/正确, B=错/不正确
        return { success: true, type: 'single', answer: isCorrect ? 'A' : 'B', explanation: '[本地题库-判断]' };
      }
    }

    // 3. 选择题：答案已经是字母格式
    if (bankAnswerArray.length > 0 && bankAnswerArray.every(a => /^[A-Z]$/i.test(a))) {
      const answer = scanType === 'single'
        ? bankAnswerArray[0].toUpperCase()
        : bankAnswerArray.map(a => a.toUpperCase());
      return { success: true, type: scanType, answer, explanation: '[本地题库]' };
    }

    // 4. 答案是文本而非字母 - 通过选项文本匹配找到对应字母
    if (bankAnswerArray.length > 0 && scannedQuestion.options && scannedQuestion.options.length > 0) {
      const matchedLabels = [];
      for (const bankAns of bankAnswerArray) {
        const norm = (bankAns || '').replace(/[^a-z0-9\u4e00-\u9fff]/gi, '').toLowerCase();
        for (const opt of scannedQuestion.options) {
          const optNorm = (opt.text || '').replace(/[^a-z0-9\u4e00-\u9fff]/gi, '').toLowerCase();
          if (optNorm === norm || optNorm.includes(norm) || norm.includes(optNorm)) {
            matchedLabels.push(opt.label.toUpperCase());
            break;
          }
        }
      }
      if (matchedLabels.length > 0) {
        const answer = scanType === 'single' ? matchedLabels[0] : matchedLabels;
        return { success: true, type: scanType, answer, explanation: '[本地题库-文本匹配]' };
      }
    }

    // 5. 兜底：直接使用原始答案
    const finalAnswer = bankAnswerArray.length === 1 ? bankAnswerArray[0] : bankAnswerArray;
    return {
      success: true,
      type: scanType,
      answer: finalAnswer,
      explanation: '[本地题库]'
    };
  }

  function getTypeLabel(type) {
    const labels = {
      single: "单选题",
      multiple: "多选题",
      fill: "填空题",
    };
    return labels[type] || type;
  }

  // Parse AI response
  function parseAIResponse(responseText) {
    // Try to extract JSON from response
    let jsonStr = responseText;

    // Handle markdown code blocks
    const jsonMatch = responseText.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (jsonMatch) {
      jsonStr = jsonMatch[1];
    }

    // Try to find JSON object
    const objMatch = jsonStr.match(/\{[\s\S]*\}/);
    if (objMatch) {
      jsonStr = objMatch[0];
    }

    jsonStr = jsonStr.trim();

    // 1. 直接解析
    try {
      return JSON.parse(jsonStr);
    } catch (e) {
      // continue to fallback
    }

    // 2. 修复常见 JSON 问题后重试
    try {
      let fixed = jsonStr
        .replace(/,\s*([}\]])/g, "$1")           // 移除尾逗号
        .replace(/'/g, '"')                       // 单引号→双引号
        .replace(/[\r\n]+/g, " ")                 // 换行→空格
        .replace(/\t/g, " ");                     // Tab→空格
      return JSON.parse(fixed);
    } catch (e) {
      // continue to fallback
    }

    // 3. 只提取 answer 字段（核心数据），跳过容易出错的 explanation
    try {
      // 匹配 "answer": "X" 或 "answer": ["A","B"]
      const answerMatch = jsonStr.match(/"answer"\s*:\s*("(?:[^"\\]|\\.)*"|\[[\s\S]*?\])/);
      const typeMatch = jsonStr.match(/"type"\s*:\s*"(single|multiple|fill)"/);
      if (answerMatch) {
        const answerValue = JSON.parse(answerMatch[1]);
        return {
          type: typeMatch ? typeMatch[1] : "single",
          answer: answerValue,
          explanation: ""
        };
      }
    } catch (e) {
      // continue to fallback
    }

    // 4. 最后兜底：直接从文本提取答案字母
    const letterMatch = jsonStr.match(/"answer"\s*:\s*"([A-Z])"/);
    if (letterMatch) {
      return { type: "single", answer: letterMatch[1], explanation: "" };
    }
    const arrayMatch = jsonStr.match(/"answer"\s*:\s*\[\s*"([A-Z])"(?:\s*,\s*"([A-Z])")*\s*\]/);
    if (arrayMatch) {
      const letters = arrayMatch[0].match(/"([A-Z])"/g).map(s => s.replace(/"/g, ""));
      return { type: "multiple", answer: letters, explanation: "" };
    }

    throw new Error("无法从AI响应中提取答案: " + jsonStr.substring(0, 100));
  }

  // Apply answer to question
  async function applyAnswer(question, answer) {
    switch (question.type) {
      case "single":
        await applySingleAnswer(question, answer);
        break;
      case "multiple":
        await applyMultipleAnswer(question, answer);
        break;
      case "fill":
        await applyFillAnswer(question, answer);
        break;
    }
  }

  // Apply single choice answer
  async function applySingleAnswer(question, answer) {
    const answerLetter = String(answer.answer).toUpperCase();

    for (const option of question.options) {
      if (option.label === answerLetter) {
        await clickElement(option.element);
        break;
      }
    }
  }

  // Apply multiple choice answer
  async function applyMultipleAnswer(question, answer) {
    let answers = answer.answer;
    if (typeof answers === "string") {
      answers = answers.split("").filter((c) => /[A-Z]/.test(c));
    }

    for (const option of question.options) {
      if (answers.includes(option.label)) {
        await clickElement(option.element);
        await sleep(200);
      }
    }
  }

  // Apply fill-in-the-blank answer
  async function applyFillAnswer(question, answer) {
    let answers = answer.answer;
    if (!Array.isArray(answers)) {
      answers = [answers];
    }

    for (let i = 0; i < Math.min(answers.length, question.inputs.length); i++) {
      const input = question.inputs[i];
      const value = String(answers[i]);

      // Focus and fill
      input.focus();
      await sleep(100);

      // Clear existing value
      input.value = "";

      // Set new value
      input.value = value;

      // Trigger events
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new Event("change", { bubbles: true }));
      input.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true }));

      await sleep(100);
    }
  }

  // Click element - 增强版，支持多种点击方式
  async function clickElement(element) {
    if (!element) {
      console.warn("[AI答题助手] clickElement: element为空");
      return;
    }

    console.log("[AI答题助手] 点击元素:", element.tagName, element.className);

    // 1. 如果是input元素（radio/checkbox）
    if (element.tagName === "INPUT") {
      const inputType = element.type?.toLowerCase();
      if (inputType === "radio" || inputType === "checkbox") {
        // 优先尝试点击关联的label（腾讯问卷需要这样）
        if (element.id) {
          const label = document.querySelector(`label[for="${element.id}"]`);
          if (label) {
            console.log("[AI答题助手] 找到关联label，点击label");
            label.click();
            await sleep(50);
            if (element.checked) {
              console.log("[AI答题助手] 通过label点击成功");
              return;
            }
          }
        }

        // 尝试点击 Vue/React 组件包装器（如 .ws-checkbox, .ws-radio 等）
        const wrapper = element.closest(
          '.ws-checkbox, .ws-radio, [role="checkbox"], [role="radio"]'
        );
        if (wrapper) {
          console.log("[AI答题助手] 找到组件包装器，点击wrapper:", wrapper.className);
          wrapper.click();
          await sleep(50);
          return;
        }

        // 无包装器时，直接操作原生 input
        element.checked = true;
      }
      element.dispatchEvent(
        new MouseEvent("click", { bubbles: true, cancelable: true })
      );
      element.dispatchEvent(new Event("change", { bubbles: true }));
      element.dispatchEvent(new Event("input", { bubbles: true }));
      return;
    }

    // 2. 查找内部的input元素
    const input = element.querySelector(
      'input[type="radio"], input[type="checkbox"]'
    );
    if (input) {
      input.checked = true;
      input.dispatchEvent(
        new MouseEvent("click", { bubbles: true, cancelable: true })
      );
      input.dispatchEvent(new Event("change", { bubbles: true }));
      input.dispatchEvent(new Event("input", { bubbles: true }));
    }

    // 3. 点击元素本身
    element.dispatchEvent(
      new MouseEvent("click", {
        bubbles: true,
        cancelable: true,
        view: window,
      })
    );

    // 4. 尝试直接调用click方法
    if (typeof element.click === "function") {
      element.click();
    }

    // 5. 对于某些框架，可能需要触发mousedown/mouseup
    element.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    await sleep(50);
    element.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
  }

  // Scroll to element
  function scrollToElement(element) {
    if (!element) return;
    element.scrollIntoView({ behavior: "smooth", block: "center" });
  }

  // Highlight element - 高亮当前正在处理的题目（绿色流动光效）
  function highlightElement(element) {
    if (!element) return;

    // 保存原始样式
    element.dataset.originalOutline = element.style.outline || "";
    element.dataset.originalOutlineOffset = element.style.outlineOffset || "";
    element.dataset.originalBoxShadow = element.style.boxShadow || "";
    element.dataset.originalPosition = element.style.position || "";

    // 确保元素有定位以便添加伪元素
    if (getComputedStyle(element).position === "static") {
      element.style.position = "relative";
    }

    // 添加流动光效样式（如果还没添加）
    if (!document.getElementById("ai-answer-highlight-style")) {
      const style = document.createElement("style");
      style.id = "ai-answer-highlight-style";
      style.textContent = `
        @keyframes ai-border-flow {
          0% { background-position: 0% 50%; }
          50% { background-position: 100% 50%; }
          100% { background-position: 0% 50%; }
        }
        .ai-answer-processing {
          position: relative !important;
        }
        .ai-answer-processing::before {
          content: '';
          position: absolute;
          top: -3px;
          left: -3px;
          right: -3px;
          bottom: -3px;
          background: linear-gradient(90deg, #3b82f6, #6366f1, #8b5cf6, #6366f1, #3b82f6);
          background-size: 300% 100%;
          border-radius: 8px;
          z-index: -1;
          animation: ai-border-flow 2s ease infinite;
        }
        .ai-answer-processing::after {
          content: '';
          position: absolute;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: white;
          border-radius: 6px;
          z-index: -1;
        }
      `;
      document.head.appendChild(style);
    }

    // 应用高亮样式
    element.style.boxShadow = "0 0 20px rgba(59, 130, 246, 0.5)";
    element.style.transition = "all 0.3s ease";
    element.style.zIndex = "1";

    // 添加动画类
    element.classList.add("ai-answer-processing");
  }

  // Remove highlight - 移除高亮
  function removeHighlight(element) {
    if (!element) return;

    // 恢复原始样式
    element.style.outline = element.dataset.originalOutline || "";
    element.style.outlineOffset = element.dataset.originalOutlineOffset || "";
    element.style.boxShadow = element.dataset.originalBoxShadow || "";
    element.style.position = element.dataset.originalPosition || "";
    element.style.zIndex = "";

    // 移除动画类
    element.classList.remove("ai-answer-processing");
  }

  // Send log to popup
  function sendLog(level, text) {
    chrome.runtime.sendMessage({ type: "log", level, text });
  }

  // Update stats in popup
  function updateStats() {
    chrome.runtime.sendMessage({
      type: "updateStats",
      questionCount: questions.length,
      answeredCount,
    });
  }

  // Send complete message
  function sendComplete() {
    chrome.runtime.sendMessage({ type: "complete", answeredCount });
  }

  // Sleep utility
  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  // Initialize
  console.log("AI自动答题助手已加载");

  // 初始化模板系统
  if (window.templateManager) {
    window.templateManager
      .init()
      .then(() => {
        console.log("模板系统初始化完成");
      })
      .catch((error) => {
        console.error("模板系统初始化失败:", error);
      });
  }
})();

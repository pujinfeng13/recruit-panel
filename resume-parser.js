/* =====================================================================
 * resume-parser.js · 简历解析模块（纯逻辑，无 DOM 依赖）
 * ---------------------------------------------------------------------
 * 供招聘面板（recruit-panel.html）解析 Word(.docx) / PDF / TXT 简历：
 *  - docx：依赖 JSZip（注入），解压 word/document.xml 提取文本
 *  - pdf ：依赖 pdfjs（window.pdfjsLib，注入），按 y/x 坐标还原行结构逐页提取
 *  - 字段抽取：四层策略（文件名 → 键值对 → 教育/工作区块 → 全文启发式）
 * 挂载：window.ResumeParser（浏览器）/ module.exports（node 测试）
 * ===================================================================== */
(function (root, factory) {
  var api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.ResumeParser = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ---------- docx：XML 文本提取（正则，不依赖 DOMParser） ---------- */
  function textFromDocxXml(xml) {
    var t = String(xml || '');
    t = t.replace(/<w:p[^>]*>/g, '\n')
         .replace(/<w:tab[^>]*\/?>/g, '\t')
         .replace(/<w:br[^>]*\/?>/g, '\n')
         .replace(/<w:cr[^>]*\/?>/g, '\n');
    t = t.replace(/<[^>]+>/g, '');
    t = t.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
         .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
         .replace(/&#(\d+);/g, function (m, n) { return String.fromCodePoint(+n); })
         .replace(/&#x([0-9a-fA-F]+);/g, function (m, n) { return String.fromCodePoint(parseInt(n, 16)); });
    return t.split('\n').map(function (s) { return s.trim(); }).filter(Boolean).join('\n');
  }

  /* ---------- docx：解压 + 提取 ---------- */
  async function parseDocx(buf, JSZip) {
    if (!JSZip) throw new Error('Word 解析库（JSZip）未加载');
    var zip = await JSZip.loadAsync(buf);
    var entry = zip.file('word/document.xml');
    if (!entry) throw new Error('非法的 .docx 文件（缺少 document.xml）');
    var xml = await entry.async('string');
    return textFromDocxXml(xml);
  }

  /* ---------- pdf.js 文本项 → 按版面还原行 ----------
   * pdf.js 的 items 是零散的文本片段，直接 join 会把整页拼成一行，
   * 导致「行级键值对 / 教育区块 / 工作区块」全部失效。
   * 这里按 y 坐标聚类（同 y 视为同行）、x 坐标排序，还原真实行结构；
   * 同一行内 x 间距过大时补空格（还原「姓名 张三」「138 1234 5678」等）。
   * 注意：pdf.js 的 transform[5] 是 PDF 页面坐标（原点在左下角、y 向上），
   * 视觉上从上到下 = y 从大到小；按升序排序会把整页文本颠倒，
   * 必须按 y 降序。 */
  function layoutLines(items) {
    var segs = [];
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      if (!it || !it.str) continue;
      var tr = it.transform || [1, 0, 0, 1, 0, 0];
      segs.push({ x: tr[4] || 0, y: tr[5] || 0, w: it.width || 0, s: it.str });
    }
    segs.sort(function (a, b) { return (b.y - a.y) || (a.x - b.x); });
    var lines = [], cur = null, curY = null;
    for (var k = 0; k < segs.length; k++) {
      var g = segs[k];
      if (curY === null || Math.abs(g.y - curY) > 4) {
        if (cur && cur.length) lines.push(cur);
        cur = [g]; curY = g.y;
      } else {
        cur.push(g);
      }
    }
    if (cur && cur.length) lines.push(cur);
    return lines.map(function (line) {
      line.sort(function (a, b) { return a.x - b.x; });
      var out = '', prev = null;
      for (var j = 0; j < line.length; j++) {
        var seg = line[j];
        if (prev && (seg.x - (prev.x + prev.w)) > 4) out += ' ';
        out += seg.s;
        prev = seg;
      }
      return out.trim();
    }).filter(Boolean);
  }

  /* BOSS 直聘导出 PDF 会在页面里铺水印追踪串（混合大小写数字的长串）及其
   * 碎片（孤立字母），会混进文本行干扰解析，这里过滤掉。 */
  function isJunkLine(s) {
    if (!s) return true;
    if (s.length >= 28 && /^[A-Za-z0-9-]+$/.test(s) && /[a-z]/.test(s) && /[A-Z]/.test(s) && /[0-9]/.test(s)) return true;
    if (s.length === 1 && /[a-zA-Z]/.test(s)) return true;
    return false;
  }

  /* ---------- pdf：pdf.js 逐页提取 ---------- */
  async function parsePdf(buf, pdfjs) {
    if (!pdfjs || typeof pdfjs.getDocument !== 'function') throw new Error('PDF 解析库（pdf.js）未加载');
    if (pdfjs.GlobalWorkerOptions && !pdfjs.GlobalWorkerOptions.workerSrc) {
      var workerUrl = '';
      // 扩展页（面板）：manifest 的 worker-src 'self' 允许 chrome-extension:// worker，直接用直链
      if (typeof location !== 'undefined' && location.protocol === 'chrome-extension:') {
        workerUrl = (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.getURL)
          ? chrome.runtime.getURL('panel/lib/pdf.worker.min.js') : '';
      } else {
        // 其他环境：尝试从页面已注入的 pdf.worker 脚本标签推断
        try {
          if (typeof document !== 'undefined') {
            var tags = document.getElementsByTagName('script');
            for (var i = 0; i < tags.length; i++) {
              if (/pdf\.worker/i.test(tags[i].src || '')) { workerUrl = tags[i].src; break; }
            }
          }
        } catch (e) {}
      }
      if (workerUrl) pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;
    }
    var data;
    if (typeof Buffer !== 'undefined' && Buffer.isBuffer(buf)) data = new Uint8Array(buf);
    else if (buf instanceof Uint8Array) data = buf;
    else data = new Uint8Array(buf);
    var task = pdfjs.getDocument({ data: data });
    var doc = await task.promise;
    var out = [];
    for (var i = 1; i <= doc.numPages; i++) {
      var page = await doc.getPage(i);
      var tc = await page.getTextContent();
      var lines = layoutLines(tc.items).filter(function (s) { return !isJunkLine(s); });
      out.push(lines.join('\n'));
      page.cleanup();
    }
    try { doc.destroy(); } catch (e) {}
    return out.filter(Boolean).join('\n');
  }

  /* =====================================================================
   * 字段抽取 · 第一层：文件名
   * 「张三-产品经理.pdf」「李四的简历.docx」「Java开发工程师-张三.pdf」…
   * ===================================================================== */
  var POS_HINT = /(工程师|经理|主管|总监|专员|顾问|设计师|开发|运营|产品|测试|前端|后端|算法|架构|讲师|教师|编辑|记者|销售|市场|人事|财务|会计|行政|助理|技术员|实习生|负责人|主任|律师|医生|护士|策划|采购|物流|客服|专家|研究员|分析师|教练|秘书|BD|QA|PM|HR|运营专员|产品助理)/i;

  function parseFileName(name) {
    var res = { name: '', position: '', exp: null, city: '' };
    var base = String(name || '').replace(/\\/g, '/').split('/').pop().replace(/\.[^.]+$/, '').trim();
    if (!base) return res;
    var clean = base.replace(/(个人)?(中文|英文)?简历|求职简历|resume|个人资料|求职信/ig, '');
    clean = clean.replace(/[-_—–·,，.。()（）【】\[\]{}<>《》"'“”‘’\s]+/g, '|')
                 .replace(/^\|+|\|+$/g, '');
    var parts = clean.split('|').map(function (s) { return s.trim(); }).filter(Boolean);
    if (!parts.length) return res;
    /* 姓名：2-4 字中文段，排除职位词/城市词/语气助词，并剥离「N年以内」等年限后缀。
       BOSS 下载文件名如「【品牌策划...】袁俪珊 1年以内.pdf」，「上海」这类城市词
       必须跳过，否则会被误判为姓名。 */
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i].replace(/(\d+(?:\.\d+)?)\s*年(?:以上|以内|以下|左右)?$/, '').trim();
      if (/^[\u4e00-\u9fa5·]{2,4}$/.test(p) && !POS_HINT.test(p) && !CITY_STOP.test(p) && !/[的了在与之和]$/.test(p)) {
        res.name = p; parts.splice(i, 1); break;
      }
    }
    if (!res.position) {
      for (var j = 0; j < parts.length; j++) {
        var q = parts[j];
        if (q && q.length <= 20 && (POS_HINT.test(q) || /^[A-Za-z][A-Za-z0-9+ ]{1,15}$/.test(q))) {
          res.position = q; break;
        }
      }
    }
    /* 工作年限：文件名自带「N年 / N年以内 / N年以上」（BOSS 下载文件名的年限标签）。
       注意排除「3-5年」这类区间（多为岗位要求年限，不是候选人年限，避免误标）。 */
    var skipExp = /(\d{1,2}(?:\.\d)?)\s*[-–—~]\s*(\d{1,2}(?:\.\d)?)\s*年/.test(base);
    if (res.exp == null && !skipExp) {
      for (var e0 = 0; e0 < parts.length; e0++) {
        var em = parts[e0].match(/^(\d{1,2}(?:\.\d)?)\s*年(?:以上|以内|以下|左右)?$/);
        if (em) { res.exp = parseFloat(em[1]); break; }
      }
    }
    /* 期望城市：baseX 优先，否则第一个城市词（BOSS 文件名如「base福州_上海」） */
    for (var c0 = 0; c0 < parts.length; c0++) {
      var bm = parts[c0].match(/^base(.{2,6})$/i);
      if (bm && CITY_STOP.test(bm[1])) { res.city = bm[1]; break; }
    }
    if (!res.city) {
      for (var c2 = 0; c2 < parts.length; c2++) {
        if (CITY_STOP.test(parts[c2])) { res.city = parts[c2]; break; }
      }
    }
    if (!res.name || !res.position) {
      var m = base.match(/^([\u4e00-\u9fa5·]{2,4})\s*[-_—–,，.]\s*([A-Za-z0-9\u4e00-\u9fa5·+（）() ]{2,20})$/);
      if (m) {
        if (!res.name) res.name = m[1];
        if (!res.position && POS_HINT.test(m[2])) res.position = m[2].trim();
      }
    }
    if (!res.name) {
      var m2 = base.match(/^([\u4e00-\u9fa5·]{2,4}?)(?:的)?(?:个人简历|求职简历|简历|resume)?$/i);
      if (m2 && !/^(个人|求职|应聘|简历|最新|中文|英文)/.test(m2[1])) res.name = m2[1];
    }
    return res;
  }

  /* =====================================================================
   * 字段抽取 · 辅助函数
   * ===================================================================== */
  var EDU_RANK = ['博士', '硕士', '本科', '大专'];

  /* 学历级别：博士4 硕士3 本科2 大专/中专/高中1 其他0 */
  function rankEdu(e) {
    var s = String(e || '').replace(/\s+/g, '');
    if (/博士/.test(s)) return 4;
    if (/硕士|研究生/.test(s)) return 3;
    if (/本科|学士/.test(s)) return 2;
    if (/大专|中专|高中/.test(s)) return 1;
    return 0;
  }

  function normEdu(e) {
    var s = String(e || '').replace(/\s+/g, '');
    if (/中专|高中/.test(s) || s === '大专') return '大专';
    if (/研究生|硕士/.test(s) || /^MBA$/i.test(s)) return '硕士';
    if (/博士/.test(s)) return '博士';
    if (/本科|学士/.test(s)) return '本科';
    return s;
  }

  /* ---------- 部首归一化 ----------
   * 部分 PDF 生成器（如 BOSS 直聘导出模板）把 CJK 常用字重编码为
   * 康熙部首（U+2F00–U+2FDF）或补充部首（U+2E80–U+2EF3），如
   * 「⼤学」「⼯作」「⼿机」「⻉泰妮」。这些码点不在 \u4e00-\u9fa5
   * 内，会导致所有中文正则失效。这里映射回标准简体/汉字。 */
  /* CJK 补充部首区 U+2E80-U+2EF3（简体部首，源自 UCD EquivalentUnifiedIdeograph） */
  var RAD_SUPP = {
    '\u2E81': '厂',
    '\u2E82': '乛',
    '\u2E83': '乚',
    '\u2E84': '乙',
    '\u2E85': '亻',
    '\u2E86': '冂',
    '\u2E87': '𠘨',
    '\u2E88': '刀',
    '\u2E89': '刂',
    '\u2E8A': '卜',
    '\u2E8B': '㔾',
    '\u2E8C': '小',
    '\u2E8D': '小',
    '\u2E8E': '兀',
    '\u2E8F': '尣',
    '\u2E90': '尢',
    '\u2E91': '𡯂',
    '\u2E92': '巳',
    '\u2E93': '幺',
    '\u2E94': '彑',
    '\u2E95': '𫜹',
    '\u2E96': '忄',
    '\u2E97': '心',
    '\u2E98': '扌',
    '\u2E99': '攵',
    '\u2E9B': '旡',
    '\u2E9C': '日',
    '\u2E9D': '月',
    '\u2E9E': '歺',
    '\u2E9F': '母',
    '\u2EA0': '民',
    '\u2EA1': '氵',
    '\u2EA2': '氺',
    '\u2EA3': '灬',
    '\u2EA4': '爫',
    '\u2EA5': '爫',
    '\u2EA6': '丬',
    '\u2EA7': '牛',
    '\u2EA8': '犭',
    '\u2EA9': '王',
    '\u2EAA': '𤴔',
    '\u2EAB': '目',
    '\u2EAC': '示',
    '\u2EAD': '礻',
    '\u2EAE': '𥫗',
    '\u2EAF': '糹',
    '\u2EB0': '纟',
    '\u2EB1': '罓',
    '\u2EB2': '罒',
    '\u2EB3': '㓁',
    '\u2EB4': '冗',
    '\u2EB5': '𦉫',
    '\u2EB6': '羊',
    '\u2EB7': '𦍌',
    '\u2EB8': '𦍋',
    '\u2EB9': '耂',
    '\u2EBA': '肀',
    '\u2EBB': '聿',
    '\u2EBC': '肉',
    '\u2EBD': '𦥑',
    '\u2EBE': '艹',
    '\u2EBF': '艹',
    '\u2EC0': '艹',
    '\u2EC1': '虎',
    '\u2EC2': '衤',
    '\u2EC3': '覀',
    '\u2EC4': '西',
    '\u2EC5': '见',
    '\u2EC6': '角',
    '\u2EC7': '𧢲',
    '\u2EC8': '讠',
    '\u2EC9': '贝',
    '\u2ECA': '𧾷',
    '\u2ECB': '车',
    '\u2ECC': '辶',
    '\u2ECD': '辶',
    '\u2ECE': '辶',
    '\u2ECF': '邑',
    '\u2ED0': '钅',
    '\u2ED1': '長',
    '\u2ED2': '镸',
    '\u2ED3': '长',
    '\u2ED4': '门',
    '\u2ED5': '𨸏',
    '\u2ED6': '阝',
    '\u2ED7': '雨',
    '\u2ED8': '青',
    '\u2ED9': '韦',
    '\u2EDA': '页',
    '\u2EDB': '风',
    '\u2EDC': '飞',
    '\u2EDD': '食',
    '\u2EDE': '𩙿',
    '\u2EDF': '飠',
    '\u2EE0': '饣',
    '\u2EE1': '𩠐',
    '\u2EE2': '马',
    '\u2EE3': '骨',
    '\u2EE4': '鬼',
    '\u2EE5': '鱼',
    '\u2EE6': '鸟',
    '\u2EE7': '卤',
    '\u2EE8': '麦',
    '\u2EE9': '黄',
    '\u2EEA': '黾',
    '\u2EEB': '斉',
    '\u2EEC': '齐',
    '\u2EED': '歯',
    '\u2EEE': '齿',
    '\u2EEF': '竜',
    '\u2EF0': '龙',
    '\u2EF1': '龜',
    '\u2EF2': '亀',
    '\u2EF3': '龟'
  };

  /* 康熙部首完整 214 序列表（U+2F00 + i → 标准汉字，兜底用） */
  var KANGXI = '一丨丶丿乙亅二亠人儿入八冂冖冫几凵刀力勹匕匚匸十卜卩厂厶又口囗土士夂夊夕大女子宀寸小尢尸屮山巛工己巾干幺广廴廾弋弓彐彡彳心戈戶手支攴文斗斤方无日曰月木欠止歹殳毋比毛氏气水火爪父爻爿片牙牛犬玄玉瓜瓦甘生用田疋疒癶白皮皿目矛矢石示禸禾穴立竹米糸缶网羊羽老而耒耳聿肉臣自至臼舌舛舟艮色艸虍虫血行衣襾見角言谷豆豕豸貝赤走足身車辛辰辵邑酉釆里金長門阜隶隹雨靑非面革韋韭音頁風飛食首香馬骨高髟鬥鬯鬲鬼魚鳥鹵鹿麥麻黃黍黑黹黽鼎鼓鼠鼻齊齒龍龜龠';


  function normRadicals(s) {
    if (!/[\u2E80-\u2FDF]/.test(s)) return s;
    var out = '';
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      var r = RAD_SUPP[s[i]];
      if (r) { out += r; continue; }
      if (c >= 0x2F00 && c < 0x2F00 + KANGXI.length) out += KANGXI[c - 0x2F00];
      else out += s[i];
    }
    return out;
  }

  function normalizeText(t) {
    /* 先做部首归一化（BOSS 导出 PDF 把汉字重编码为康熙部首，如「⼤学」「⼿机」），
       再做全角/标点归一化 */
    return normRadicals(String(t || ''))
      .replace(/\r/g, '')
      .replace(/：/g, ':')
      .replace(/（/g, '(').replace(/）/g, ')')
      .replace(/　/g, ' ')
      .replace(/；/g, ';')
      .replace(/，/g, ',');
  }

  /* 简历头部「个人信息」高频标签：专业/学校抽取时必须排除 */
  var INFO_STOP = /^(政治面貌|中共党员|共青团员|预备党员|群众|健康状况|婚姻状况|婚姻|籍贯|民族|身高|体重|兴趣爱好|个人爱好|户口|现居住地|居住地|期望薪资|到岗时间|政治面貌|政治|面貌|党员|团员|证件号码|身份证号|微信|邮箱|电话|手机|出生年月|年龄|性别|学历|毕业院校|学校|专业|工作年限|工作经验|期望城市|求职意向|应聘职位|目前所在公司|现居地|星座|血型)$/;

  /* 常见城市（文件名中易与姓名混淆的 2 字城市） */
  var CITY_STOP = /^(北京|上海|广州|深圳|杭州|成都|重庆|武汉|南京|天津|苏州|西安|长沙|青岛|大连|厦门|福州|宁波|无锡|合肥|昆明|哈尔滨|济南|郑州|东莞|佛山|沈阳|石家庄|南昌|贵阳|兰州|太原|长春|常州|徐州|南通|温州|嘉兴|绍兴|金华|台州|泉州|珠海|惠州|中山|海口|三亚|南宁|桂林|乌鲁木齐|呼和浩特|银川|西宁|拉萨|香港|澳门|台湾|雄安|廊坊|保定|烟台|潍坊|威海|淄博|临沂|济宁|洛阳|宜昌|襄阳|株洲|湘潭|岳阳|衡阳|桂林|柳州|贵阳|绵阳|德阳|南充|宜宾|泸州|自贡|芜湖|蚌埠|马鞍山|安庆|阜阳|滁州|九江|赣州|上饶|泉州|漳州|莆田|宁德|龙岩|珠海|汕头|江门|湛江|茂名|肇庆|惠州|梅州|揭阳|潮州|柳州|北海|防城港|钦州|玉林|贵阳|遵义|安顺|昆明|曲靖|玉溪|兰州|天水|西宁|银川|乌鲁木齐|克拉玛依|呼和浩特|包头|鄂尔多斯|拉萨|日喀则|哈尔滨|齐齐哈尔|大庆|牡丹江|长春|吉林|四平|沈阳|大连|鞍山|抚顺|本溪|丹东|锦州|营口|盘锦|铁岭|葫芦岛|太原|大同|阳泉|长治|晋城|朔州|运城|临汾|吕梁|西安|铜川|宝鸡|咸阳|渭南|延安|汉中|安康|商洛|石家庄|唐山|秦皇岛|邯郸|邢台|保定|张家口|承德|沧州|廊坊|衡水|合肥|芜湖|蚌埠|淮南|马鞍山|淮北|铜陵|安庆|黄山|滁州|阜阳|宿州|六安|亳州|池州|宣城)$/;

  /* 值清洗：去空白、按冒号截断（同行多标签时取第一段）、去尾部括号内容 */
  function cleanVal(v) {
    var s = String(v == null ? '' : v).trim().split(/[:：]/)[0].trim();
    s = s.replace(/\s+/g, '');
    return s;
  }

  function isSchoolName(s) {
    return /(大学|学院|学校|分校|University|College|Institute|Academy)$/i.test(s) && s.length >= 3 && s.length <= 32;
  }

  function cleanSchool(s) {
    return String(s || '')
      .replace(/^(毕业于|就读于|就读|毕业|来自|于|在|现|曾)+/, '')
      .replace(/^(博士|硕士|本科|学士|大专|中专|研究生|全日制|非全日制|统招|专升本|在读)?\s*(学位|学历)?\s*[|｜·、\s]*/, '')
      .replace(/(本科|硕士|博士|大专|研究生|学历|教育|学习|学校|院校|专业|全日制)+$/, '');
  }

  var MAJOR_DICT = [
    '计算机科学与技术', '软件工程', '信息管理与信息系统', '电子信息工程', '通信工程',
    '机械设计制造及其自动化', '电气工程及其自动化', '数据科学与大数据技术', '国际经济与贸易',
    '人力资源管理', '工商管理', '市场营销', '行政管理', '物流管理', '电子商务', '旅游管理',
    '会计学', '金融学', '经济学', '法学', '汉语言文学', '新闻学', '广告学', '教育学', '心理学',
    '自动化', '机械工程', '车辆工程', '土木工程', '建筑学', '工程管理', '物联网工程', '网络工程',
    '人工智能', '信息安全', '统计学', '数学与应用数学', '信息与计算科学', '物理学', '应用物理学',
    '化学', '应用化学', '生物科学', '生物技术', '药学', '临床医学', '护理学', '英语', '日语',
    '应用心理学', '学前教育', '小学教育', '视觉传达设计', '环境设计', '产品设计', '工业设计',
    '数字媒体艺术', '材料科学与工程', '能源与动力工程', '公共事业管理', '社会学', '历史学',
    '环境科学', '食品科学与工程', '服装设计', '动画', '供应链管理', '大数据管理与应用',
    '传播学', '公共关系学', '网络与新媒体', '广播电视学', '数字媒体技术', '品牌传播',
    '品牌管理', '会展经济与管理', '社会工作', '行政管理', '文化产业管理', '新闻传播学'
  ];

  /* 区块标题识别（行首，允许带编号） */
  var SEC_TITLES = ['教育经历', '教育背景', '教育情况', '教育信息', '学习经历', '最高学历', '教育',
    '工作经历', '工作经验', '工作背景', '工作履历', '从业经历', '职业经历', '主要工作经历',
    '正式工作', '核心实习', '工作实习', '全职经历',
    '基本信息', '个人信息', '个人资料', '基本资料', '个人概况', '联系方式', '联系信息',
    '自我评价', '专业技能', '项目经验', '项目经历', '证书', '获奖', '荣誉', '语言能力',
    '兴趣爱好', '个人爱好', '校园经历', '社团经历', '社会实践', '实习经历', '培训经历',
    '个人优势', '自我描述', '自我简介', '个人总结'];

  function stripSectionTitle(line) {
    var s = line.replace(/^[\d一二三四五六七八九十、.．(（)\s]+/, '');
    for (var i = 0; i < SEC_TITLES.length; i++) {
      if (s.indexOf(SEC_TITLES[i]) === 0) { s = s.slice(SEC_TITLES[i].length); break; }
    }
    return s.replace(/^[:：\s|｜]+/, '').trim();
  }

  function detectSection(line) {
    var head = line.replace(/^[\d一二三四五六七八九十、.．(（)\s]+/, '');
    if (/^(教育经历|教育背景|教育情况|教育信息|学习经历|最高学历|教育|教育及培训)/.test(head)) return 'edu';
    /* v1.2.16：带「（年）」或冒号后仅跟数字的行是表单键值（如「工作经验（年）」「工作年限：0」），
       不是区块标题——若按标题剥离，残段（如「（年）」）会污染工作区块的公司识别。
       注意「工作经历：」「工作经历：2019.06-至今 …」是正常标题，不在此列。 */
    if (/^(工作经历|工作经验|工作背景|工作履历|从业经历|职业经历|主要工作经历|工作及项目|工作\/项目|正式工作|核心实习|工作实习|全职经历)/.test(head)
        && !/[（(]\s*年\s*[）)]/.test(head) && !/[:：]\s*\d{1,3}\s*年?\s*$/.test(head)) return 'work';
    if (/^(基本信息|个人信息|个人资料|基本资料|个人概况|联系方式|联系信息)/.test(head)) return 'info';
    if (/^(自我评价|专业技能|项目经验|项目经历|证书|获奖|荣誉|语言能力|兴趣爱好|个人爱好|校园经历|社团经历|社会实践|实习经历|培训经历|个人优势|自我描述|自我简介|个人总结)/.test(head)) return 'other';
    return '';
  }

  /* 行级键值对扫描 */
  function kvScan(line, out, nowY) {
    var m;
    /* 姓名 */
    if (!out.name) {
      m = line.match(/^\s*(?:姓\s*名|name)\s*[:：]\s*(.{1,20})$/i);
      if (m) {
        var nm = cleanVal(m[1]);
        if (/^[\u4e00-\u9fa5·]{2,4}$/.test(nm) && !/[的了在与之和]$/.test(nm)) out.name = nm;
      }
    }
    /* 年龄：显式年龄 / N岁（容忍「年 龄」中间空格） */
    if (out.age == null) {
      m = line.match(/(?:年\s*龄|age)\s*[:：]?\s*(\d{1,2})\s*(?:岁)?/i);
      if (m && +m[1] >= 16 && +m[1] <= 70) out.age = +m[1];
      else {
        m = line.match(/(\d{1,2})\s*岁(?!\d)/);
        if (m && +m[1] >= 16 && +m[1] <= 70) out.age = +m[1];
      }
    }
    /* 出生年月 → 年龄 */
    if (out.age == null) {
      m = line.match(/(?:出生年月|出生日期|出生时间|生日|出生|birth)[^:：]*[:：]?\s*((?:19|20)\d{2})/i);
      if (m) out.age = Math.max(0, nowY - +m[1]);
    }
    /* 个人信息行（含性别标记）中的裸年份 → 出生年（如「男 | 1993年5月 | 8年经验」） */
    if (out.age == null && /[男女]/.test(line) && !/(入职|毕业|教育|入学|项目|经历|公司|至今|学校|大学|学院)/.test(line)) {
      m = line.match(/((?:19|20)\d{2})\s*年?\s*\d{0,2}\s*月?/);
      if (m && +m[1] >= 1950 && +m[1] <= nowY) out.age = Math.max(0, nowY - +m[1]);
    }
    /* 学历（容忍「本 科」「大 专」等内部空格） */
    if (!out.edu) {
      m = line.match(/(?:最高学历|统招学历|全日制学历|学\s*历|学\s*位)\s*[:：]?\s*(博士|硕士|研究生|本\s*科|大\s*专|中\s*专|高\s*中|学士)/);
      if (m) out.edu = normEdu(m[1]);
    }
    /* 学校 */
    if (!out.school) {
      m = line.match(/(?:毕业院校|毕业学校|毕业大学|最高学历院校|学校名称|所在院校|院校名称|学校)\s*[:：]\s*([A-Za-z\u4e00-\u9fa5（）()· ]{2,40})/);
      if (m) {
        var sch = String(m[1]).trim().replace(/\s+/g, ' ').replace(/[*#_~`]/g, '');
        if (isSchoolName(sch)) out.school = sch;
      }
    }
    /* 专业（排除「政治面貌」「中共党员」等个人信息词） */
    if (!out.major) {
      m = line.match(/(?:所学专业|主修专业|专业名称|专\s*业|专业方向)\s*[:：]\s*([^\n,;]{1,24})/);
      if (m) {
        var maj = cleanVal(m[1]).replace(/[()（）]*(本科|硕士|博士|大专|研究生)[()（）]*$/, '');
        if (maj && maj.length <= 24 && !INFO_STOP.test(maj)) out.major = maj;
      }
    }
    /* 工作经验：年限标签 / 经验标签 / N年经验 / 纯数字+年 */
    if (out.exp == null) {
      m = line.match(/(?:工作年限|工作年数|从业年限|从业年数|工作经验年限|工作经验|经验年限|年限)\s*(?:[（(]\s*年\s*[）)])?\s*[:：]?\s*(\d{1,2}(?:\.\d)?)\s*年?/);
      if (!m) m = line.match(/(?:工作经验|从业经验|工作经历)\s*[:：]?\s*(\d{1,2}(?:\.\d)?)\s*年/);
      if (!m) m = line.match(/(\d{1,2}(?:\.\d)?)\s*年(?:以上|多)?(?:工作|从业|行业|相关)?经验/);
      if (!m) m = line.match(/^(\d{1,2}(?:\.\d)?)\s*年/);
      if (m) out.exp = parseFloat(m[1]);
    }
    /* 经验：从业/入职时间「X年-至今」跨度、工作N年、N+年、裸「经验：N年」 */
    if (out.exp == null) {
      m = line.match(/(?:从业时间|入职时间|参加工作(?:时间)?|工作起始时间|工作开始时间|在职时间|工作时间)\s*[:：]?\s*((?:19|20)\d{2})[./年-]*\s*\d{0,2}\s*月?\s*[-–—~至到]\s*(?:至今|现在|今)/);
      if (m && +m[1] >= 1970 && +m[1] <= nowY) out.exp = Math.max(0, nowY - +m[1]);
    }
    if (out.exp == null) {
      m = line.match(/工作(\d{1,2}(?:\.\d)?)\s*年/);
      if (!m) m = line.match(/(\d{1,2}(?:\.\d)?)\s*\+\s*年/);
      if (!m) m = line.match(/(?:经验)\s*[:：]?\s*(\d{1,2}(?:\.\d)?)\s*年/);
      if (m) out.exp = parseFloat(m[1]);
    }
    /* 期望城市 */
    if (!out.city) {
      m = line.match(/(?:期望城市|期望工作地|意向城市|意向工作地|期望工作地点|期望工作城市|工作地点|工作城市|现居住地|居住地|所在城市|现居地|现居)\s*[:：]\s*([^\n,;]{1,12})/);
      if (m) {
        var c = cleanVal(m[1]);
        if (!/^(不限|无|任意|全国|都可|均可|面议)$/.test(c)) out.city = c;
      }
    }
    /* 电话（标签式，容忍空格分隔） */
    if (!out.phone) {
      m = line.match(/(?:手\s*机|电\s*话|联系电话|联系\s*方式|手机号码|mobile|phone|tel)[^:：]*[:：]\s*([0-9\s\-（）()]{7,20})/i);
      if (m) {
        var ph = m[1].replace(/[^\d]/g, '');
        if (/^1[3-9]\d{9}$/.test(ph)) out.phone = ph;
      }
    }
    /* 当前公司 */
    if (!out.currentCompany) {
      m = line.match(/(?:目前所在公司|当前公司|现任职公司|现公司|在职公司|现任公司|目前公司|现就职|现工作单位)\s*[:：]\s*([A-Za-z0-9\u4e00-\u9fa5（）()·]{2,30})/);
      if (m) {
        var comp = cleanVal(m[1]);
        if (comp.length >= 2) out.currentCompany = comp;
      }
    }
    /* 应聘/期望职位 */
    if (!out.position) {
      m = line.match(/(?:求职意向|应聘职位|期望职位|意向岗位|求职岗位|目标岗位|应聘岗位|期望岗位|求职职位|意向职位|应聘方向|期望职业)\s*[:：]\s*([^\n,;]{2,24})/);
      if (m) {
        var pos = String(m[1]).trim().replace(/\s+/g, ' ');
        if (pos.length >= 2) out.position = pos;
      }
    }
    /* 毕业年份 */
    if (out.gradYear == null) {
      m = line.match(/(?:毕业时间|毕业年份|毕业年度|毕业日期)\s*[:：]?\s*((?:19|20)\d{2})/);
      if (m && +m[1] >= 1970 && +m[1] <= nowY + 5) out.gradYear = +m[1];
    }
  }

  /* ---------- 教育区块解析 ---------- */
  function majorFromLine(line, school) {
    var s = String(line || '');
    s = s.replace(/((?:19|20)\d{2})[./年-]*\s*\d{0,2}\s*[-–—~至到]\s*((?:19|20)\d{2}|至今|现在|今)[^\s]*/g, ' ');
    s = s.replace(/((?:19|20)\d{2})\s*年/g, ' ');
    if (school) {
      s = s.split(school).join(' ');
      /* 校名被拆行时本行可能只有不带「大学」的前半段，也一并剥掉 */
      s = s.split(school.replace(/(大学|学院|分校)$/, '')).join(' ');
    }
    s = s.replace(/(博士|硕士|研究生|本科|大专|中专|高中|学士|学位|全日制|非全日制|统招|专升本)/g, ' ');
    s = s.replace(/\([^()]*\)|（[^（）]*）/g, ' ');
    s = s.replace(/[^\u4e00-\u9fa5A-Za-z0-9+·]+/g, ' ').trim();
    if (!s) return '';
    var parts = s.split(/\s+/);
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i];
      if (p.length >= 2 && p.length <= 16 && !INFO_STOP.test(p) && !/^(大学|学院|分校|学校)$/.test(p)) {
        return p.replace(/(专业|方向)$/, '');
      }
    }
    return '';
  }

  function parseEduBlock(lines, nowY) {
    var res = { school: '', major: '', edu: '', gradYear: null };
    if (!lines || !lines.length) return res;
    var i, m;
    /* 最高学历行窗口：BOSS 模板常把「大学」拆成独立文本段（基线不同被分行），
       若存在硕士/博士行，则以它为首行取 4 行窗口统一扫描，避免先命中
       低学历（本科）的学校/专业/毕业年份；无硕士行则退回前 4 行。 */
    var mi = -1;
    for (i = 0; i < lines.length; i++) {
      if (/(硕士|博士|研究生)/.test(lines[i])) { mi = i; break; }
    }
    var win = mi >= 0 ? lines.slice(mi, Math.min(lines.length, mi + 4)) : lines.slice(0, 4);
    var winText = win.join(' ');
    /* 学校：贪婪匹配完整校名（「英国南安普顿 大学」「华中科技大学武昌分校」）。 */
    m = winText.match(/([\u4e00-\u9fa5]{2,16}\s*(?:大学|学院|分校))/);
    if (m) {
      var sch = cleanSchool(m[1]).replace(/\s+/g, '');
      if (isSchoolName(sch)) res.school = sch;
    }
    if (!res.school) {
      m = winText.match(/([A-Za-z]{3,40}?(?:University|College|Institute|Academy))/i);
      if (m) res.school = m[1];
    }
    if (!res.school) {
      /* 英文校名在前：「University of Melbourne」 */
      m = winText.match(/((?:The\s+)?(?:University|College|Institute|Academy)\s+of\s+[A-Za-z ]{2,24})/i);
      if (m) res.school = m[1].trim();
    }
    /* 学历：优先硕士/博士（最高学历），其次窗口内本科等 */
    for (i = 0; i < lines.length; i++) {
      m = lines[i].match(/(博士|硕士|研究生)/);
      if (m) { res.edu = normEdu(m[1]); break; }
    }
    if (!res.edu) {
      for (i = 0; i < win.length; i++) {
        m = win[i].match(/(本科|学士|大专|中专|高中)/);
        if (m) { res.edu = normEdu(m[1]); break; }
      }
    }
    /* 毕业年份：窗口内时间范围结束年 / N届 */
    for (i = 0; i < win.length; i++) {
      m = win[i].match(/((?:19|20)\d{2})[./年-]*\s*\d{0,2}\s*[-–—~至到]\s*((?:19|20)\d{2})/);
      if (m && +m[2] >= 1970 && +m[2] <= nowY + 5) { res.gradYear = +m[2]; break; }
      m = win[i].match(/((?:19|20)\d{2})\s*届/);
      if (m) { res.gradYear = +m[1]; break; }
    }
    /* 专业：窗口内词典/行内词优先（「信息管理学院 电子商务 本科」取「电子商务」而非学院名） */
    for (i = 0; i < win.length; i++) {
      var maj = '';
      for (var d0 = 0; d0 < MAJOR_DICT.length; d0++) {
        if (MAJOR_DICT[d0].length >= 3 && win[i].indexOf(MAJOR_DICT[d0]) !== -1) { maj = MAJOR_DICT[d0]; break; }
      }
      if (!maj) maj = majorFromLine(win[i], res.school);
      if (maj) { res.major = maj; break; }
    }
    return res;
  }

  /* 时间跨度（月）：「2022.07-2025.09」「2026/03-至今」→ 月数；不匹配返回 null */
  function spanMonths(line, nowY) {
    var m = line.match(/((?:19|20)\d{2})[./年-]*\s*(\d{1,2})?\s*月?\s*[-–—~至到]\s*((?:19|20)\d{2})[./年-]*\s*(\d{1,2})?\s*月?/);
    var y1, m1, y2, m2;
    if (m) {
      y1 = +m[1]; m1 = +(m[2] || 1); y2 = +m[3]; m2 = +(m[4] || 1);
    } else {
      m = line.match(/((?:19|20)\d{2})[./年-]*\s*(\d{1,2})?\s*月?\s*[-–—~至到]\s*(?:至今|现在|今)/);
      if (!m) return null;
      y1 = +m[1]; m1 = +(m[2] || 1); y2 = nowY; m2 = new Date().getMonth() + 1;
    }
    if (y1 < 1970 || y1 > nowY || y2 < y1 || (y2 === y1 && m2 < m1)) return null;
    return (y2 * 12 + m2) - (y1 * 12 + m1);
  }

  /* ---------- 工作区块解析：当前公司 / 职位 / 经验（时间跨度） ---------- */
  function parseWorkBlock(lines, nowY) {
    var res = { company: '', position: '', exp: null };
    if (!lines || !lines.length) return res;
    var i, m, line;
    /* 当前公司：优先含「至今/现在」的工作行（最新工作），排除教育行；
       否则取第一行 */
    var compLine = null;
    for (i = 0; i < lines.length; i++) {
      if (/(至今|现在|今)/.test(lines[i]) && !/(大学|学院|学位|学历|硕士|博士|本科|大专|专业|教育|学校)/.test(lines[i])) { compLine = lines[i]; break; }
    }
    if (!compLine) compLine = lines[0];
    /* ① 完整法定后缀（有限公司/有限责任公司）——先匹配整串，避免被「科技」等短词截断 */
    m = compLine.match(/([\u4e00-\u9fa5A-Za-z0-9（）()·]{2,30}?(?:有限公司|有限责任公司))/);
    if (!m) {
      /* ② 简称后缀（公司/集团/科技/银行…） */
      m = compLine.match(/([\u4e00-\u9fa5A-Za-z0-9（）()·]{2,24}?(?:公司|集团|科技|网络|软件|信息|银行|证券|保险|医院|事务所|工作室|传媒|文化|咨询|商贸|实业|制造|电子|研究院|设计院))/);
    }
    if (m) {
      var comp = cleanVal(m[1]);
      if (comp.length >= 2) res.company = comp;
    }
    if (!res.company) {
      var tm = compLine.match(/((?:19|20)\d{2})[./年-]*\s*\d{0,2}/);
      var head = tm ? compLine.slice(0, tm.index) : compLine;
      var segs = head.split(/[\s|｜·、,，]+/).filter(Boolean);
      for (i = 0; i < segs.length; i++) {
        var c = segs[i];
        /* v1.2.14b：含冒号/以括号开头的残段（标签剥离产物，如「（年）：0」）不是公司名 */
        if (c.length >= 2 && c.length <= 30 && !/^\d/.test(c) && !/[:：]/.test(c) && !/^[（(【[]/.test(c) && !POS_HINT.test(c)) { res.company = c; break; }
      }
    }
    /* 职位：工作行内「公司名 职位 时间」结构（对整行匹配，公司名不含职位词，风险低） */
    var m2 = compLine.match(/([\u4e00-\u9fa5A-Za-z0-9+·]{2,14}?(?:工程师|经理|主管|总监|专员|顾问|设计师|开发|运营|测试|助理|讲师|教师|编辑|记者|销售|市场|人事|财务|会计|行政|策划|采购|客服|法务|技术员|负责人|专家|研究员|分析师|秘书|HR|BD|QA|PM))/);
    if (m2) res.position = m2[1];
    /* 经验：全部工作行时间跨度的最大值（按月差算，向上取整，与 BOSS 年限标签一致） */
    var maxSpan = null;
    for (i = 0; i < lines.length; i++) {
      line = lines[i];
      var sm = spanMonths(line, nowY);
      if (sm != null && sm >= 0) {
        if (maxSpan == null || sm > maxSpan) maxSpan = sm;
        continue;
      }
      var m4 = line.match(/((?:19|20)\d{2})\s*年.*(?:入职|参加工作|加入|开始)/);
      if (m4 && +m4[1] >= 1970 && +m4[1] <= nowY) {
        var sp2 = Math.max(0, nowY - +m4[1]) * 12;
        if (maxSpan == null || sp2 > maxSpan) maxSpan = sp2;
      }
    }
    if (maxSpan != null) res.exp = Math.max(1, Math.ceil(maxSpan / 12));
    return res;
  }

  /* =====================================================================
   * 字段抽取 · 主入口（四层）
   * ===================================================================== */
  function extractFields(text, fileName) {
    var t = normalizeText(text);
    var lines = t.split('\n').map(function (s) { return s.trim(); }).filter(Boolean);
    var nowY = new Date().getFullYear();
    var out = { name: '', position: '', edu: '', school: '', major: '', exp: null, city: '', age: null, phone: '', email: '', gradYear: null, currentCompany: '' };
    var m;

    /* v1.2.14：BOSS 简历 PDF 表单常见「标签与值分两行、无冒号」排版（如「学历\n大专」「毕业年份\n2002」）。
       先把「标签行 + 下一行取值」合并为「标签：值」，即可复用下方 kvScan（原本只认同行冒号）。
       占位提示（如「自动识别/可改」「如 Java 后端工程师」「请填写」），以及「值本身又是另一个标签」的情况跳过。 */
    (function mergeFormPairs() {
      var MERGE_LABEL = /^(姓名|性别|年龄|手机|电话|邮箱|微信|期望城市|期望薪资|求职意向|应聘职位|当前公司|最高学历|学历|学位|毕业院校|毕业学校|所在院校|院校名称|学校名称|学校|所学专业|主修专业|专业名称|专业|毕业年份|毕业时间|毕业日期|工作年限|工作经验|从业年限|经验年限|工作年数)/;
      /* v1.2.18b：MERGE_LABEL 形如 /^(姓名|…)/（无 $ 结尾），剥掉「^（」后只余尾部「)」，
         /\)\$$/ 匹配不到，须再补剥一次裸「)」，否则拼接 (?:…) 产生不配对括号 */
      var MERGE_BODY = MERGE_LABEL.source.replace(/^\^\(/, '').replace(/\)\$$/, '').replace(/\)$/, '');
      /* 经验类标签合并后规范化为「工作年限」（原样保留「工作经验：0」会被工作区块标题识别吃掉，
         剥离标题后残段「（年）：0」还会被当成公司名） */
      var EXP_LABEL = /^(工作年限|工作经验|从业年限|从业年数|工作年数|经验年限|年限)$/;
      var out2 = [];
      for (var i = 0; i < lines.length; i++) {
        var ln = lines[i];
        var bare = ln.replace(/[*:：]/g, '').replace(/[ \t]*（[^）]*）[ \t]*/g, '').replace(/[ \t]*\([^)]*\)[ \t]*/g, '').trim();
        if (MERGE_LABEL.test(bare) && !/[:：]/.test(ln) && i + 1 < lines.length) {
          var v = lines[i + 1].trim();
          var vBare = v.replace(/[*:：]/g, '').replace(/[ \t]*（[^）]*）[ \t]*/g, '').replace(/[ \t]*\([^)]*\)[ \t]*/g, '').trim();
          var vHead = v.split(/[:：]/)[0].replace(/[*\s]/g, '');
          var isLabel = MERGE_LABEL.test(vBare) || MERGE_LABEL.test(vHead)
            || /^(自我评价|专业技能|项目经验|教育经历|工作经历|基本信息|联系方式|政治面貌)$/.test(vBare);
          /* v1.2.16：值若是工作条目（年份开头/含至今），说明当前行是区块标题而非表单标签，不合并。
             例外：毕业年份/时间/日期的值本身就是年份，不做此判断 */
          var isYearValLabel = /^(毕业年份|毕业时间|毕业日期|毕业年度)$/.test(bare);
          if (v && v.length <= 30 && !isLabel && !/自动识别|可改|请填写|示例|例：|点击填写|未填写|不限|^如/.test(v)
              && (isYearValLabel || (!/^(?:19|20)\d{2}/.test(v) && !/至今/.test(v)))) {
            out2.push((EXP_LABEL.test(bare) ? '工作年限' : bare) + '：' + v);
            i++;
            continue;
          }
        }
        /* v1.2.18：同行「标签 值」空格分隔形式（BOSS 表单渲染为「学校 清华大学」「年龄 25 岁」），
           合并为「标签：值」供 kvScan 复用。要求标签后紧跟空白，避免「专业从事…」词首误伤。 */
        if (!/[:：]/.test(ln)) {
          var mSL = ln.match(new RegExp('^(?:' + MERGE_BODY + ')(?:（[^）]*）)?[\\s]+(.+)$'));
          if (mSL && mSL[1] && mSL[1].trim()) {
            var vSL = mSL[1].trim();
            if (vSL.length <= 30) {
              var vSLBare = vSL.replace(/[*:：]/g, '').replace(/[ \t]*（[^）]*）[ \t]*/g, '').replace(/[ \t]*\([^)]*\)[ \t]*/g, '').trim();
              var isSLabel = MERGE_LABEL.test(vSLBare) || MERGE_LABEL.test(vSL.split(/[:：]/)[0].replace(/[*\s]/g, ''))
                || /^(自我评价|专业技能|项目经验|教育经历|工作经历|基本信息|联系方式|政治面貌)$/.test(vSLBare);
              if (!isSLabel && !/自动识别|可改|请填写|示例|例：|点击填写|未填写|不限|^如/.test(vSL)) {
                var tagPart = ln.slice(0, ln.length - vSL.length).replace(/[*:：\s]/g, '').replace(/[ \t]*（[^）]*）[ \t]*/g, '').replace(/[ \t]*\([^)]*\)[ \t]*/g, '');
                out2.push((EXP_LABEL.test(tagPart) ? '工作年限' : tagPart) + '：' + vSL);
                continue;
              }
            }
          }
        }
        out2.push(ln);
      }
      lines = out2;
    })();

    /* —— 第一层：文件名 —— */
    var fn = parseFileName(fileName || '');
    out.name = fn.name;
    out.position = fn.position;

    /* —— 第二层：全文级（电话 / 邮箱） —— */
    m = t.match(/1[3-9]\d{9}(?!\d)/);
    if (m) out.phone = m[0];
    var emailRe = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, em, emBest = '';
    while ((em = emailRe.exec(t))) { if (em[0].length > emBest.length) emBest = em[0]; }
    if (emBest) out.email = emBest;

    /* —— 第三层：行级扫描（键值对 + 区块跟踪） —— */
    var section = '';
    var eduLines = [], workLines = [], infoLines = [];
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      var sec = detectSection(line);
      if (sec) {
        section = sec;
        var body = stripSectionTitle(line);
        if (body) {
          if (section === 'edu') eduLines.push(body);
          else if (section === 'work') workLines.push(body);
          else if (section === 'info') infoLines.push(body);
          kvScan(body, out, nowY);
        }
        continue;
      }
      if (section === 'edu') eduLines.push(line);
      else if (section === 'work') workLines.push(line);
      else if (section === 'info') infoLines.push(line);
      kvScan(line, out, nowY);
    }

    /* —— 第三层B：浓缩信息行启发 ——
     * 覆盖「男 | 1995.06 | 本科 | 6年经验」「1996.08 · 硕士 · 男」这类
     * 无标签的竖线/圆点分隔简历头部行，抽取学历与出生年份（→年龄）。 */
    var compactLines = infoLines.concat(lines.slice(0, 8));
    for (var ci = 0; ci < compactLines.length; ci++) {
      var cl = compactLines[ci];
      if (!cl || cl.length > 40) continue;
      /* 学历：浓缩行中的独立学历词 */
      if (!out.edu && !/(院校|学校|教育|专业|经历|毕业院校|毕业学校)/.test(cl)) {
        m = cl.match(/(博士|硕士|研究生|本科|大专|中专|高中)/);
        if (m) out.edu = normEdu(m[1]);
      }
      /* 年龄：孤立的出生年份（前后须为分隔符/行首行尾，如「1995.06」「1993 年 5 月」），
         排除教育/工作行的时间跨度（「2014.09 - 2018.06」两侧年份均会因年龄范围校验被拒） */
      if (out.age == null && !/(教育|经历|项目|至今|公司|院校|毕业|微信|qq|QQ|账号|邮箱|邮件|电话|手机|@)/.test(cl)) {
        m = cl.match(/(?:^|[|｜·\s])((?:19|20)\d{2})\s*[./年]\s*\d{0,2}\s*月?(?:$|[|｜·\s])/);
        if (m) {
          var a = Math.max(0, nowY - +m[1]);
          if (a >= 16 && a <= 70) out.age = a;
        }
      }
    }

    /* —— 第四层A：区块兜底 —— */
    var edu = parseEduBlock(eduLines, nowY);
    if (!out.school) out.school = edu.school;
    if (!out.major) out.major = edu.major;
    if (!out.edu) out.edu = edu.edu;
    else if (edu.edu && rankEdu(edu.edu) > rankEdu(out.edu)) out.edu = edu.edu; /* 教育区块级别更高则覆盖（如基本信息写本科、教育经历有硕士） */
    if (out.gradYear == null) out.gradYear = edu.gradYear;
    var work = parseWorkBlock(workLines, nowY);
    if (!out.currentCompany) out.currentCompany = work.company;
    if (out.exp == null) out.exp = work.exp;
    if (!out.position && work.position) out.position = work.position;
    /* 文件名兜底：工作年限 / 期望城市（文本显式信息优先于文件名） */
    if (out.exp == null && fn.exp != null) out.exp = fn.exp;
    if (!out.city && fn.city) out.city = fn.city;

    /* —— 第四层B：全文启发式 —— */
    /* 姓名：前几行纯中文段（支持「袁俪珊 (Lisa)」这类带英文昵称的页头） */
    if (!out.name) {
      for (var n = 0; n < lines.length && n < 8; n++) {
        var line0 = lines[n];
        var compact = line0.replace(/\s+/g, '');
        var mN = line0.match(/^([\u4e00-\u9fa5·]{2,4})[\s(（]*[A-Za-z·\s]*(?:\)|）)?$/);
        var cand = mN ? mN[1] : (/^[\u4e00-\u9fa5·]{2,4}$/.test(compact) ? compact : '');
        if (cand &&
            !/^(个人简历|简历|求职|应聘|联系方式|基本信息|教育|工作|自我|技能|项目|姓名|电话|邮箱|地址|期望|目标|性别|年龄|学历|学校|专业|居住|手机|毕业|求职意向)/.test(cand) &&
            !/[的了在与之和]$/.test(cand)) {
          out.name = cand; break;
        }
      }
    }
    /* 学历：全文最高 */
    if (!out.edu) {
      for (var r = 0; r < EDU_RANK.length; r++) {
        if (t.indexOf(EDU_RANK[r]) !== -1) { out.edu = EDU_RANK[r]; break; }
      }
      if (!out.edu && /研究生/.test(t)) out.edu = '硕士';
    }
    /* 学校：全文第一个大学/学院（贪婪匹配完整校名，排除「第十一届全国大学生广告艺术大赛」这类干扰） */
    if (!out.school) {
      var reSch = /([\u4e00-\u9fa5]{2,16}\s*(?:大学|学院|分校))/g, mm;
      while ((mm = reSch.exec(t))) {
        var cand = cleanSchool(mm[1]).replace(/\s+/g, '');
        var ctx = t.slice(Math.max(0, mm.index - 2), mm.index + mm[1].length + 2);
        if (isSchoolName(cand) &&
            !/期|间|在校|毕业(后|于)|学历|就读/.test(cand) &&
            !/(学生|大赛|竞赛|同学|大学生|创意|比赛|运动会|课堂|毕业典礼|学生会)/.test(cand) &&
            !/(大学期间|大学生活|大学生|大赛|全国|第[一二三四五六七八九十\d]+届)/.test(ctx) &&
            !/^(全国|第[一二三四五六七八九十\d]+届|所在|在校|就读)/.test(cand)) {
          out.school = cand; break;
        }
      }
      if (!out.school) {
        var reSch2 = /([A-Za-z]{3,40}?(?:University|College|Institute|Academy))|((?:The\s+)?(?:University|College|Institute|Academy)\s+of\s+[A-Za-z ]{2,24})/i;
        m = t.match(reSch2);
        if (m) out.school = (m[1] || m[2] || '').trim();
      }
    }
    /* 专业：常见专业词典（长词直接采纳；短词需专业上下文） */
    if (!out.major) {
      for (var d = 0; d < MAJOR_DICT.length; d++) {
        var kw = MAJOR_DICT[d], idx = t.indexOf(kw);
        if (idx === -1) continue;
        var ctx2 = t.slice(Math.max(0, idx - 6), idx + kw.length + 6);
        if ((kw.length >= 6 && !/专员|经理|主管|总监|负责|担任/.test(ctx2)) || /专业|主修|就读|学习|本科|硕士|博士/.test(ctx2)) {
          out.major = kw; break;
        }
      }
    }
    /* 经验：毕业年份反推 */
    if (out.exp == null && out.gradYear) out.exp = Math.max(0, nowY - out.gradYear);

    return out;
  }

  /* ---------- 编排：文件 → 文本 + 字段 ---------- */
  var RESUME_EXTS = ['docx', 'pdf', 'txt', 'md', 'json'];
  async function parseResume(file, deps) {
    var name = file.name || 'resume';
    var ext = name.toLowerCase().split('.').pop() || '';
    if (RESUME_EXTS.indexOf(ext) === -1) throw new Error('仅支持 Word(.docx) / PDF / TXT 文件');
    var text = '';
    if (ext === 'docx') text = await parseDocx(await file.arrayBuffer(), deps.JSZip);
    else if (ext === 'pdf') text = await parsePdf(await file.arrayBuffer(), deps.pdfjs);
    else if (ext === 'json') {
      try { text = JSON.stringify(JSON.parse(await file.text())); } catch (e) { throw new Error('JSON 简历解析失败：' + e.message); }
    }
    else text = await file.text();
    if (!text || !text.trim()) throw new Error('未能从文件中提取到文本（可能为扫描件图片 PDF，请使用 Word 版）');
    return { fileName: name, text: text, fields: extractFields(text, name) };
  }

  /* ---------- Excel / CSV 导入（招聘面板批量导入候选人表格） ----------
   * 无需外部库：CSV 纯文本解析；xlsx 借助已注入的 JSZip 解压 OOXML（zip + sharedStrings + worksheet XML）。
   * 通过表头（首行）按中英文同义词映射到候选人字段；无姓名的整行跳过。 */
  var EXCEL_FIELD_SYN = {
    name:           ['姓名','名字','name'],
    position:       ['应聘职位','职位','岗位','意向职位','期望职位','求职意向','position','job'],
    edu:            ['学历','教育程度','教育','edu','degree'],
    school:         ['学校','毕业院校','院校','毕业学校','school','university'],
    major:          ['专业','主修','专业方向','major'],
    exp:            ['经验','工作经验','工作年限','年限','经验(年)','exp','experience','years'],
    age:            ['年龄','岁','age'],
    gradYear:       ['毕业年份','毕业时间','毕业','gradyear','graduation'],
    city:           ['期望城市','城市','工作地','意向城市','工作城市','city','location'],
    phone:          ['联系电话','电话','手机','手机号码','手机号','phone','tel','mobile'],
    email:          ['邮箱','电子邮件','email','mail'],
    currentCompany: ['当前公司','公司','现公司','在职公司','company','employer'],
    source:         ['来源','渠道','source'],
    status:         ['状态','阶段','status','stage'],
    note:           ['备注','备注说明','说明','note','comment'],
    salaryText:     ['薪资','期望薪资','salary','pay'],
    gender:         ['性别','gender','sex'],
    uploadAt:       ['上传时间','创建时间','导入时间','uploaded','upload'],
    updatedAt:      ['最近更新','更新时间','updated']
  };

  function colToIndex(letters){
    var idx = 0;
    for (var i = 0; i < letters.length; i++) idx = idx * 26 + (letters.charCodeAt(i) - 64);
    return idx - 1;
  }
  /* xlsx 采用正则直接抽取 OOXML 片段（sharedStrings / worksheet），无需 DOMParser，浏览器与 node 测试环境均可运行 */
  function toNum(s){ if (s == null || s === '') return null; var n = parseFloat(String(s).replace(/[^\d.\-]/g, '')); return isNaN(n) ? null : n; }
  function toInt(s){ if (s == null || s === '') return null; var n = parseInt(String(s).replace(/[^\d\-]/g, ''), 10); return isNaN(n) ? null : n; }
  function toDate(s){
    if (!s) return null;
    var n = parseFloat(s);
    if (!isNaN(n) && n > 20000 && n < 80000) return (n - 25569) * 86400 * 1000; // Excel 序列日期
    var t = Date.parse(s); return isNaN(t) ? null : t;
  }

  /* 纯文本 CSV → 二维数组（支持引号包裹、字段内逗号/换行、双引号转义、BOM、CRLF/LF） */
  function parseCSVText(text){
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    var rows = [], row = [], field = '', i = 0, n = text.length, inQ = false;
    while (i < n) {
      var ch = text[i];
      if (inQ) {
        if (ch === '"') { if (text[i+1] === '"') { field += '"'; i += 2; continue; } inQ = false; i++; continue; }
        field += ch; i++; continue;
      }
      if (ch === '"') { inQ = true; i++; continue; }
      if (ch === ',') { row.push(field); field = ''; i++; continue; }
      if (ch === '\r') { i++; continue; }
      if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }
      field += ch; i++;
    }
    if (field.length || row.length) { row.push(field); rows.push(row); }
    return rows;
  }

  /* xlsx 缓冲区 → 二维数组（取第一个工作表）。仅依赖 JSZip，用正则直接抽取 OOXML 片段（无需 DOMParser，浏览器/node 通吃） */
  async function parseXlsxBuffer(buffer, deps){
    var JSZip = deps && deps.JSZip;
    if (!JSZip) throw new Error('JSZip 未加载，无法解析 xlsx');
    var zip = await JSZip.loadAsync(buffer);
    var shared = [];
    var ssFile = zip.file('xl/sharedStrings.xml') || zip.file('xl/sharedStrings2.xml');
    if (ssFile) {
      var ssXml = await ssFile.async('string');
      var siRe = /<si>([\s\S]*?)<\/si>/g, tRe = /<t[^>]*>([\s\S]*?)<\/t>/g, m, tm;
      while ((m = siRe.exec(ssXml))) {
        var siInner = m[1], s = '', t;
        while ((t = tRe.exec(siInner))) s += t[1];
        shared.push(s);
      }
    }
    var sheets = Object.keys(zip.files).filter(function (p) { return /^xl\/worksheets\/sheet\d+\.xml$/i.test(p); }).sort();
    if (!sheets.length) throw new Error('xlsx 中未找到工作表');
    var sheetXml = await zip.file(sheets[0]).async('string');
    var grid = [];
    var rowRe = /<row\b[^>]*>([\s\S]*?)<\/row>/g, rm;
    while ((rm = rowRe.exec(sheetXml))) {
      var rowInner = rm[1];
      var tmp = {}, maxCol = -1;
      var cellRe = /<c\b([^>]*)>([\s\S]*?)<\/c>/g, cm;
      while ((cm = cellRe.exec(rowInner))) {
        var attrs = cm[1], inner = cm[2];
        var rM = attrs.match(/\br="([^"]*)"/), tM = attrs.match(/\bt="([^"]*)"/);
        var ref = rM ? rM[1] : '';
        var col = colToIndex(ref.replace(/[0-9]/g, ''));
        var t = tM ? tM[1] : '';
        var val = '';
        if (t === 's') {
          var vM = inner.match(/<v>([\s\S]*?)<\/v>/);
          var idx = vM ? parseInt(vM[1], 10) : -1;
          val = (idx >= 0 && idx < shared.length) ? shared[idx] : '';
        } else if (t === 'inlineStr' || t === 'str') {
          var isM = inner.match(/<is>([\s\S]*?)<\/is>/);
          if (isM) { var it = isM[1].match(/<t[^>]*>([\s\S]*?)<\/t>/); val = it ? it[1] : ''; }
          else { var vv = inner.match(/<v>([\s\S]*?)<\/v>/); val = vv ? vv[1] : ''; }
        } else {
          var vn = inner.match(/<v>([\s\S]*?)<\/v>/); val = vn ? vn[1] : '';
        }
        tmp[col] = val;
        if (col > maxCol) maxCol = col;
      }
      var arr = [];
      for (var cc = 0; cc <= maxCol; cc++) arr.push(tmp[cc] != null ? tmp[cc] : '');
      grid.push(arr);
    }
    return grid;
  }

  /* 二维数组（首行为表头）→ 候选人对象数组。deps.mapStatus(name) 把状态文本转为阶段 id */
  function mapGridToCandidates(grid, deps){
    deps = deps || {};
    if (!grid || !grid.length) return [];
    var header = grid[0].map(function (h) { return String(h == null ? '' : h).trim(); });
    var colMap = {};
    Object.keys(EXCEL_FIELD_SYN).forEach(function (field) {
      var syns = EXCEL_FIELD_SYN[field];
      for (var c = 0; c < header.length && colMap[field] == null; c++) {
        var h = header[c].toLowerCase();
        for (var s = 0; s < syns.length; s++) { if (h.indexOf(syns[s].toLowerCase()) !== -1) { colMap[field] = c; break; } }
      }
    });
    if (colMap.name == null && header.length && header.every(function (h) { return h === ''; })) colMap.name = 0;
    var mapStatus = deps.mapStatus || function (n) { return n; };
    var keys = Object.keys(colMap), out = [];
    for (var r = 1; r < grid.length; r++) {
      var row = grid[r]; if (!row) continue;
      var raw = {};
      keys.forEach(function (field) { var v = row[colMap[field]]; raw[field] = (v == null ? '' : String(v).trim()); });
      if (!raw.name) continue;
      var c = { name: raw.name };
      if (raw.position) c.position = raw.position;
      if (raw.edu) c.edu = raw.edu;
      if (raw.school) c.school = raw.school;
      if (raw.major) c.major = raw.major;
      if (raw.city) c.city = raw.city;
      if (raw.phone) c.phone = raw.phone;
      if (raw.email) c.email = raw.email;
      if (raw.currentCompany) c.currentCompany = raw.currentCompany;
      if (raw.source) c.source = raw.source;
      if (raw.note) c.note = raw.note;
      if (raw.salaryText) c.salaryText = raw.salaryText;
      if (raw.gender) c.gender = raw.gender;
      var exp = toNum(raw.exp); if (exp != null) c.exp = exp;
      var age = toInt(raw.age); if (age != null) c.age = age;
      var gy = toInt(raw.gradYear); if (gy != null) c.gradYear = gy;
      if (raw.status) c.status = mapStatus(raw.status);
      var up = toDate(raw.uploadAt); if (up) c.uploadAt = up;
      var ud = toDate(raw.updatedAt); if (ud) c.updatedAt = ud;
      out.push(c);
    }
    return out;
  }

  async function parseExcel(file, deps){
    deps = deps || {};
    var name = file.name || 'data';
    var ext = name.toLowerCase().split('.').pop() || '';
    if (ext === 'csv') return mapGridToCandidates(parseCSVText(await file.text()), deps);
    if (ext === 'xlsx' || ext === 'xlsm') return mapGridToCandidates(await parseXlsxBuffer(await file.arrayBuffer(), deps), deps);
    if (ext === 'xls') throw new Error('暂不支持旧版 .xls 二进制格式，请用 Excel 另存为 .xlsx 或导出 CSV 后再导入');
    throw new Error('仅支持 .xlsx / .csv 文件');
  }

  return {
    textFromDocxXml: textFromDocxXml,
    parseDocx: parseDocx,
    parsePdf: parsePdf,
    extractFields: extractFields,
    parseFileName: parseFileName,
    parseResume: parseResume,
    layoutLines: layoutLines,
    EDU_RANK: EDU_RANK,
    parseExcel: parseExcel,
    parseCSVText: parseCSVText,
    parseXlsxBuffer: parseXlsxBuffer,
    mapGridToCandidates: mapGridToCandidates
  };
});

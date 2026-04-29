const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType, ShadingType,
} = require('docx');

const NODE_MODULES = '/Users/apple/.nvm/versions/node/v22.14.0/lib/node_modules';

const NAVY = "1E3A5F";
const RED = "C0392B";
const GREY = "555555";
const LIGHT_BG = "F4F7FB";
const BORDER_GRAY = "CCCCCC";

const FONT = "PingFang SC";

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before ?? 0, after: opts.after ?? 80, line: 320 },
    alignment: opts.align ?? AlignmentType.LEFT,
    children: Array.isArray(text)
      ? text
      : [new TextRun({ text, font: FONT, size: opts.size ?? 22, bold: opts.bold, color: opts.color })]
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    border: { bottom: { color: NAVY, space: 4, style: BorderStyle.SINGLE, size: 12 } },
    children: [new TextRun({ text, font: FONT, size: 30, bold: true, color: NAVY })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 100 },
    children: [new TextRun({ text, font: FONT, size: 26, bold: true, color: NAVY })]
  });
}

function bullet(text, level = 0) {
  const children = [];
  if (Array.isArray(text)) {
    children.push(...text);
  } else {
    children.push(new TextRun({ text, font: FONT, size: 22 }));
  }
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { before: 0, after: 60, line: 320 },
    children
  });
}

function tr(text) {
  return new TextRun({ text, font: FONT, size: 22 });
}
function trB(text, color) {
  return new TextRun({ text, font: FONT, size: 22, bold: true, color });
}
function trC(text, color) {
  return new TextRun({ text, font: FONT, size: 22, color });
}

function box(paragraphs, fillColor = LIGHT_BG) {
  const border = { style: BorderStyle.SINGLE, size: 4, color: NAVY };
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: 9026, type: WidthType.DXA },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
        shading: { fill: fillColor, type: ShadingType.CLEAR },
        borders: { left: { style: BorderStyle.SINGLE, size: 24, color: NAVY }, top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
        children: paragraphs
      })]
    })]
  });
}

function tableSimple(rows, widths) {
  const border = { style: BorderStyle.SINGLE, size: 4, color: BORDER_GRAY };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((row, ri) => new TableRow({
      children: row.map((cell, ci) => new TableCell({
        width: { size: widths[ci], type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 140, right: 140 },
        shading: ri === 0 ? { fill: NAVY, type: ShadingType.CLEAR } : { fill: ci === 0 ? LIGHT_BG : "FFFFFF", type: ShadingType.CLEAR },
        borders,
        children: Array.isArray(cell)
          ? cell
          : [new Paragraph({
              spacing: { before: 0, after: 0, line: 300 },
              children: [new TextRun({
                text: cell,
                font: FONT,
                size: 22,
                bold: ri === 0,
                color: ri === 0 ? "FFFFFF" : (ci === 0 ? NAVY : "2C3E50")
              })]
            })]
      }))
    }))
  });
}

const doc = new Document({
  creator: "Claude",
  title: "代树坤简历优化建议",
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } }
  },
  numbering: {
    config: [{
      reference: "bullets",
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: "▸", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 240 } }, run: { font: FONT, color: NAVY } } },
        { level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 240 } }, run: { font: FONT } } }
      ]
    }]
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 } } },
    children: [

      // 标题
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: "代树坤简历 · 优化建议与待补充清单", font: FONT, size: 40, bold: true, color: NAVY })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 240 },
        border: { bottom: { color: NAVY, space: 8, style: BorderStyle.SINGLE, size: 12 } },
        children: [new TextRun({ text: "—— 拟用于求职：基因编辑研发 / 水产生物育种 / 分子生物学方向", font: FONT, size: 22, italics: true, color: GREY })]
      }),

      // 说明
      box([
        p([
          trB("说明：", NAVY),
          tr("本文档基于现有简历内容做的结构化分析，整体认为「专业方向清晰、科研产出过硬」，已是应届硕士中较有竞争力的版本。下文列出"),
          trB("需要本人确认或补充", RED),
          tr("的内容（主要是数字化指标），以及建议的写法示例，便于针对性修改。")
        ], { after: 0 })
      ]),

      // ----- 一、整体评价 -----
      h1("一、整体评价"),

      tableSimple([
        ["评估维度", "评价"],
        ["专业稀缺性", "「水生动物医学 + 海洋生物技术」复合背景，在水产基因编辑/育种赛道属稀缺人才"],
        ["技术深度", "完整掌握 CRISPR/Cas9 + 电穿孔 + sgRNA 设计 + 体外转录 + qRT-PCR + 原位杂交全链条"],
        ["科研产出", "Aquaculture（SCI）+《水产科学》（中文核心），对应届生属于较硬的成果"],
        ["独立工作能力", "多次「独立完成」——文献调研、方案设计、实验、论文，符合 R&D 岗位画像"],
        ["项目思维", "两个课题都明确「行业痛点 → 解决方案」，体现项目落地思维而非纯实验工"]
      ], [1800, 7226]),

      // ----- 二、已修正的问题 -----
      h1("二、已在 PDF 中修正的问题"),

      p("以下问题在新版简历 PDF 中已直接处理，无需再补充：", { after: 120 }),

      bullet([trB("错别字修正：", NAVY), tr("GraphPrism → "), trB("GraphPad Prism", RED), tr("；Alphafold3 → "), trB("AlphaFold3", RED), tr("；Snapgene → "), trB("SnapGene", RED)]),
      bullet([trB("重复词删除：", NAVY), tr("「实验技能」中"), trC("「原位杂交」", RED), tr("原本写了两次，已删除一处")]),
      bullet([trB("CET-4 分数处理：", NAVY), tr("470 分刚过线，已隐去分数，仅保留「可熟练阅读英文 SCI 文献」")]),
      bullet([trB("结构调整：", NAVY), tr("「自我介绍」前置至顶部、「论文发表」紧跟科研经历，更符合 HR 阅读习惯")]),
      bullet([trB("新增求职方向：", NAVY), tr("顶部加一行「基因编辑研发 / 水产生物育种 / 分子生物学」，让 HR 一眼定位匹配度")]),
      bullet([trB("视觉优化：", NAVY), tr("统一深蓝主色 + 红色点缀，板块用色块分隔，单页 A4 排版清爽")]),

      // ----- 三、待补充内容 -----
      h1("三、需要补充确认的内容（重要）"),

      p("简历整体最大的短板是 ", { after: 60 }),
      box([
        p([
          trB("缺乏量化指标。", RED),
          tr("当前科研经历用「攻克难题」「大幅缩短时间」等定性描述，HR 与技术面试官更关心 "),
          trB("具体的数字结果。", NAVY),
          tr("以下表格请逐项确认或补充——这是简历从「良好」升级到「出色」的关键。")
        ], { after: 0 })
      ]),

      h2("3.1  硕士课题（CRISPR/Cas9 基因编辑）"),

      tableSimple([
        ["待补充指标", "建议写法示例"],
        ["电穿孔条件参数", "电压 _____V、脉冲 _____ms、脉冲次数 _____ 次（最优组合）"],
        ["细胞/胚胎存活率", "优化后存活率从 ___% 提升至 ___%"],
        ["编辑效率（敲除/敲入）", "目标基因编辑效率达 ___%（T7E1 或测序验证）"],
        ["sgRNA 数量", "针对 ___ 个目标基因，设计并筛选 ___ 条 sgRNA"],
        ["实验样本量", "累计完成 ___ 批次电转实验，处理胚胎/细胞 ___ 枚"],
        ["对比基线", "相比传统显微注射法效率提升 ___ 倍 / 成本降低 ___%"]
      ], [3200, 5826]),

      h2("3.2  本科课题（LAMP 快速检测）"),

      tableSimple([
        ["待补充指标", "建议写法示例"],
        ["检测时间", "检测时间从传统 PCR 的 ___ 小时缩短至 ___ 分钟"],
        ["灵敏度阈值", "最低检测限达 ___ CFU/mL（或 ___ 拷贝/μL）"],
        ["特异性验证菌株数", "用 ___ 株近缘菌验证特异性，无交叉反应"],
        ["反应温度/条件", "等温扩增温度 ___℃、反应时间 ___ min"],
        ["实际应用样本数", "应用于 ___ 份养殖水样/病鱼样本，符合率 ___%"]
      ], [3200, 5826]),

      h2("3.3  论文信息补全"),

      p("两篇论文目前只写了期刊和作者位次，建议补全以下信息：", { after: 100 }),

      bullet([trB("Aquaculture 论文：", NAVY), tr("发表年份、卷期、起止页码、DOI、")
        , trB("当年影响因子（IF）", RED), tr("。Aquaculture 当前 IF 约 4.5+，建议明确写出")]),
      bullet([trB("水产科学论文：", NAVY), tr("发表年份、卷期、页码")]),
      bullet([trB("作者位次格式：", NAVY), tr("「第四作者」可保留中文，也可写成「4th author」标准格式")]),

      h2("3.4  其他可补充加分项"),

      tableSimple([
        ["项", "说明"],
        ["奖学金 / 荣誉", "国奖、学业奖学金、优秀毕业生、三好学生等（如有）"],
        ["专利", "是否有发明专利或实用新型（即使是参与人也可写）"],
        ["学术会议", "是否在中国海洋大学/中国水产学会等会议做过口头/壁报报告"],
        ["实习/横向课题", "是否参与过企业横向项目或实习（如海大集团、通威等）"],
        ["导师 / 课题组", "硕导姓名 + 课题组方向（如已征得导师同意可写）"],
        ["GPA / 排名", "本科或硕士成绩排名（如 Top 10% 等）"],
        ["其他技能", "Python / R / Linux / 生信分析（RNA-seq、ChIP-seq）等"]
      ], [2200, 6826]),

      // ----- 四、表达优化建议 -----
      h1("四、文字表达优化建议"),

      h2("4.1  「研究内容」段落的写法升级"),

      p([trB("当前写法（偏过程描述）：", GREY)]),
      box([p("熟练操作电转仪、荧光显微镜等核心仪器，攻克侏儒蛤电穿孔存活率与编辑效率平衡的技术难题。", { after: 0 })], "F8F8F8"),

      p([trB("建议写法（强调结果与价值）：", NAVY)], { before: 80 }),
      box([p("通过 L9(3⁴) 正交试验筛选最优电穿孔条件（电压 XXX V、脉冲 XXX ms），将存活率由 XX% 提升至 XX%，同时实现 XX% 的基因编辑效率，相比显微注射法成本降低约 XX%，已被课题组作为标准操作流程（SOP）采纳。", { after: 0 })], LIGHT_BG),

      h2("4.2  自我介绍可更具针对性"),
      bullet("如果目标行业明确（如水产企业 vs 科研院所 vs 生物科技公司），可针对性突出不同的关键词"),
      bullet("水产企业方向：突出「低成本、易推广、可现场检测」"),
      bullet("科研院所/读博方向：突出「文献调研、独立设计实验、论文发表、SOP 建立」"),
      bullet("生物科技公司：突出「全流程能力、问题解决、技术落地」"),

      // ----- 五、求职方向建议 -----
      h1("五、推荐投递的求职方向"),

      tableSimple([
        ["方向", "代表单位/企业", "匹配度"],
        ["水产/海洋生物科技公司 R&D", "海大集团、通威股份、国联水产、好当家", "★★★★★"],
        ["基因编辑/合成生物学公司", "齐禾生科、中科蓝晶、辉大基因（部分有水产线）", "★★★★"],
        ["科研院所", "中科院海洋所、黄海水产研究所、中国水产科学院", "★★★★★"],
        ["农业生物育种企业", "先正达、登海种业、隆平高科（水产板块）", "★★★★"],
        ["医疗诊断 / IVD 企业", "凭 LAMP 检测背景投递分子诊断研发岗", "★★★"],
        ["读博深造", "海大、厦大、中科院体系（产出已达申博水平）", "★★★★★"]
      ], [2400, 4626, 2000]),

      // ----- 六、最终行动清单 -----
      h1("六、给小代的行动清单"),

      p([trB("Step 1 ", RED), tr("（最重要）：填写第三章 3.1 / 3.2 表格中的具体数字，发回给我，我帮你直接更新到简历 PDF 中。")]),
      p([trB("Step 2 ", RED), tr("：补全两篇论文的发表年份、卷期、DOI 与影响因子。")]),
      p([trB("Step 3 ", RED), tr("：回想本硕期间是否有奖学金、专利、会议报告、实习经历等加分项，列出来。")]),
      p([trB("Step 4 ", RED), tr("：明确主投方向（企业/院所/读博），便于针对性微调自我介绍措辞。")]),
      p([trB("Step 5 ", RED), tr("：所有材料确认后，输出最终版 PDF + 一份英文版简历（如需投外企或申博）。")]),

      new Paragraph({ spacing: { before: 480 }, children: [
        new TextRun({ text: "—— 文档结束 ——", font: FONT, size: 20, italics: true, color: GREY })
      ], alignment: AlignmentType.CENTER })
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/Users/apple/Desktop/resume/代树坤简历_优化建议.docx", buf);
  console.log("docx written");
});

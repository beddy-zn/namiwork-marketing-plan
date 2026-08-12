const docx = require('docx');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType, BorderStyle, PageBreak, Footer, Header, PageNumber, ImageRun } = docx;
const fs = require('fs');
const path = require('path');

const logoBuffer = fs.readFileSync(path.join(__dirname, 'hex-logo.png'));

const BLUE = '1B3A6B', BLUE_DARK = '0F2847', BLUE_LIGHT = 'E8F0FE', GOLD = 'D4A020';
const TEXT = '1A1A2E', TEXT2 = '4A5568', TEXT3 = '8A94A6', BG_SOFT = 'F7F9FC', BORDER = 'D1D9E6';

function p(text, opts={}) {
  const {bold, color, size, italic, align, spacing} = opts;
  return new Paragraph({
    spacing:{after:spacing||200,line:360},alignment:align||AlignmentType.LEFT,
    children:[new TextRun({text,bold:bold||false,color:color||TEXT2,size:size||22,italics:italic||false})]
  });
}
function h1(t){return new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{before:400,after:200},children:[new TextRun({text:t,bold:true,color:BLUE_DARK,size:36})]})}
function h2(t){return new Paragraph({heading:HeadingLevel.HEADING_2,spacing:{before:360,after:160},children:[new TextRun({text:t,bold:true,color:BLUE,size:30})]})}
function h3(t){return new Paragraph({heading:HeadingLevel.HEADING_3,spacing:{before:280,after:120},children:[new TextRun({text:t,bold:true,color:TEXT,size:26})]})}
function h4(t){return new Paragraph({heading:HeadingLevel.HEADING_4,spacing:{before:200,after:80},children:[new TextRun({text:t,bold:true,color:TEXT,size:22})]})}
function divider(){return new Paragraph({spacing:{before:100,after:200},border:{bottom:{color:BLUE,size:6,style:BorderStyle.SINGLE,space:1}},children:[]})}
function spacer(a=200){return new Paragraph({spacing:{after:a},children:[]})}
function bullet(text){return new Paragraph({spacing:{after:80,line:340},bullet:{level:0},children:[new TextRun({text,color:TEXT2,size:21})]})}

function insightBox(label, paras){
  return new Table({width:{size:100,type:WidthType.PERCENTAGE},columnWidths:[9000],rows:[
    new TableRow({children:[new TableCell({
      width:{size:100,type:WidthType.PERCENTAGE},
      shading:{type:ShadingType.CLEAR,fill:BLUE_LIGHT},
      margins:{top:200,bottom:200,left:300,right:300},
      children:[
        new Paragraph({spacing:{after:100},children:[new TextRun({text:label,bold:true,color:GOLD,size:18})]}),
        ...paras.map(t=>new Paragraph({spacing:{after:100,line:340},children:[new TextRun({text:t,color:TEXT2,size:21})]}))
      ]
    })]})
  ]});
}
function calloutBox(label,text){
  return new Table({width:{size:100,type:WidthType.PERCENTAGE},columnWidths:[9000],rows:[
    new TableRow({children:[new TableCell({
      width:{size:100,type:WidthType.PERCENTAGE},
      shading:{type:ShadingType.CLEAR,fill:'FFF9E6'},
      margins:{top:200,bottom:200,left:300,right:300},
      borders:{left:{style:BorderStyle.SINGLE,size:12,color:'E0B000'},top:{style:BorderStyle.SINGLE,size:2,color:'F0E0A0'},bottom:{style:BorderStyle.SINGLE,size:2,color:'F0E0A0'},right:{style:BorderStyle.SINGLE,size:2,color:'F0E0A0'}},
      children:[
        new Paragraph({spacing:{after:60},children:[new TextRun({text:label,bold:true,color:'8A6D00',size:18})]}),
        new Paragraph({spacing:{line:340},children:[new TextRun({text,color:TEXT2,size:21})]})
      ]
    })]})
  ]});
}
function makeTable(headers,rows){
  const cw=Math.floor(9000/headers.length);
  const trs=[];
  trs.push(new TableRow({tableHeader:true,children:headers.map(h=>new TableCell({
    width:{size:cw,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill:BLUE_LIGHT},
    margins:{top:100,bottom:100,left:120,right:120},
    children:[new Paragraph({children:[new TextRun({text:h,bold:true,color:BLUE,size:18})]})]
  }))}));
  rows.forEach(r=>trs.push(new TableRow({children:r.map(c=>new TableCell({
    width:{size:cw,type:WidthType.DXA},margins:{top:100,bottom:100,left:120,right:120},
    children:[new Paragraph({children:[new TextRun({text:c,color:TEXT2,size:18})]})]
  }))})));
  return new Table({width:{size:100,type:WidthType.PERCENTAGE},columnWidths:Array(headers.length).fill(cw),rows:trs});
}

const doc=new Document({
  styles:{default:{document:{run:{font:'Microsoft YaHei',size:22},paragraph:{spacing:{line:360}}}}},
  background:{color:'F2F6FC'},
  sections:[{
    properties:{page:{margin:{top:1440,bottom:1440,left:1440,right:1440}}},
    headers:{default:new Header({children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[
      new ImageRun({data:logoBuffer,transformation:{width:20,height:20},type:'png'}),
      new TextRun({text:'  纳米Work 企业版',color:TEXT3,size:16})
    ]})]})},
    footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[
      new TextRun({text:'— ',color:TEXT3,size:16}),
      new TextRun({children:[PageNumber.CURRENT],color:TEXT3,size:16}),
      new TextRun({text:' —',color:TEXT3,size:16})
    ]})]})},
    children:[
      // COVER
      new Paragraph({spacing:{before:1600,after:200},alignment:AlignmentType.CENTER,children:[
        new ImageRun({data:logoBuffer,transformation:{width:80,height:80},type:'png'})
      ]}),
      new Paragraph({spacing:{after:200},alignment:AlignmentType.CENTER,children:[new TextRun({text:'纳米Work 企业版',bold:true,color:BLUE_DARK,size:22})]}),
      new Paragraph({spacing:{after:300},alignment:AlignmentType.CENTER,children:[new TextRun({text:'GTM Strategy · 2026',color:BLUE,size:20,bold:true})]}),
      new Paragraph({spacing:{after:300},alignment:AlignmentType.CENTER,children:[new TextRun({text:'360纳米Work企业版',color:TEXT,size:56,bold:true})]}),
      new Paragraph({spacing:{after:600},alignment:AlignmentType.CENTER,children:[new TextRun({text:'6个月GTM营销方案',color:BLUE,size:36,bold:true})]}),
      new Paragraph({spacing:{after:100},alignment:AlignmentType.CENTER,children:[new TextRun({text:'从品类定义到标杆落地',color:TEXT2,size:24})]}),
      new Paragraph({spacing:{after:400},alignment:AlignmentType.CENTER,children:[new TextRun({text:'不卖工具，卖用得起的AI落地服务',color:TEXT3,size:20,italics:true})]}),
      new Paragraph({spacing:{after:100},alignment:AlignmentType.CENTER,children:[new TextRun({text:'6个月周期 · 3大北极星指标 · 6场活动 · 6渠道矩阵',color:TEXT3,size:18})]}),
      new Paragraph({spacing:{after:100},alignment:AlignmentType.CENTER,children:[new TextRun({text:'中小企业 50-500人',color:TEXT3,size:18})]}),
      new Paragraph({children:[new PageBreak()]}),

      // ===== 核心战略判断 =====
      h1('核心战略判断'),
      divider(),
      insightBox('◆ 战略锚点', [
        '钉钉、飞书、企微的AI能力，本质是在既有办公生态上叠加智能体。护城河不是AI，是组织已经在里面办公——通讯录、审批流、文件、历史沉淀都在。正面进攻等于要求客户先离开一个每天都在用的系统，这是营销解决不了的问题。',
        '品类重新定义：不在"AI办公智能体"赛道和BAT竞争，在"企业AI落地服务"赛道，对手是收费几十万的咨询公司和SI集成商。从"卖工具"切换到"卖用得起的AI落地服务"。',
        '叙事框架：超级个体→超级组织。不卖功能，卖进化。',
      ]),
      spacer(),

      // ===== 一、市场判断与品类重新定义 =====
      h1('一、市场判断与品类重新定义'),
      divider(),

      h3('竞品三家模式 — 我们不打'),
      p('钉钉→千问办公、飞书→Aily、企业微信→WorkBuddy，三家BAT都是"办公生态+嵌入AI智能体"模式。用户被各自生态锁定。无生态绑定反而是我们的优势——企业不用担心用一个AI就要换一套办公系统。'),
      p('其他竞品：WPS灵犀(金山) | TRAE Work(字节) | Microsoft 365 Copilot | ChatGPT | 腾讯元宝 | 豆包 | Kimi | Notion AI'),

      h3('360纳米Work企业版 — 护城河'),

      h4('安全 = 信任底座'),
      p('不是功能卖点，而是准入门槛。"大模型出错是说错话，智能体出错是干错事" — 360纳米Work企业版原生安全、出厂内置。360累计捕获60个境外APT组织，3800+安全专家团队，覆盖2万家政企客户。'),
      p('63.6%的AI供应商未在法律文件中披露第三方数据转发（DataGrail 2026报告）。72%企业AI Agent渗透率但安全没跟上。国安部已披露三类AI泄密典型案件。安全不是"后贴的补丁"，是20年攻防积累的原生能力。'),

      h4('交付 = 服务承诺'),
      p('竞品交付"你用AI做出来的"，不一定能用。360纳米Work交付"AI替你做出来的"。端到端任务交付而非辅助生成。纳米1314任务引擎：用户输入目标→AI拆解任务→调度多专家→云端7×24执行→交付成品。'),

      h4('交付场景分级（关键设计）'),
      makeTable(
        ['级别','定义','代表场景','交付承诺'],
        [
          ['A级','可全自动交付','招投标监控、多语种内容生成、合规台账','7×24无人值守'],
          ['B级','AI交付+人工审核','合同审查、营销方案生成、客户分析','AI完成80%，人工判断'],
          ['C级','AI辅助+人主导','战略规划、创意决策','不做交付承诺'],
        ]
      ),
      spacer(100),
      p('不能笼统说"什么都能交付"。A级场景定义输入→AI执行→输出格式→质量标准→交付周期，营销只宣传A级案例。'),

      h4('FDE = 核心产能'),
      p('不是售后支持，是核心商业模式。依托360庞大白帽子群体→FDE转型+城市合伙人=网格式AI落地服务网络。FDE三级认证：训练工匠→AI业务专家→AI咨询架构师。'),

      calloutBox('💡 精妙点子：AI数字员工入职','企业不是在"买软件"，是在"招AI数字员工"。1亿Token=试用期工资360付；FDE=入职培训师；500+专家=AI员工简历库；城市合伙人=人才中介。企业对"招人"是刚需，对"买软件"有抵触。'),
      spacer(),

      // ===== 二、目标客户与优先级 =====
      h1('二、目标客户与优先级'),
      divider(),

      h3('把安全客户变成AI客户'),
      p('已付费、已信任、已有采购关系与对接人的客户——这是360最大的存量资产。动作：按三条件筛出首批300家满足：'),
      bullet('员工50-500人（中小型企业）'),
      bullet('近12个月有续费'),
      bullet('属三个Beachhead行业'),
      p('由安全侧客户经理带FDE二次拜访，不做产品推销，做免费「AI场景诊断」。从"你在卖AI工具"变成"你在帮我查AI安全漏洞"——心智完全不同。'),

      h3('三个Beachhead行业推荐'),
      makeTable(
        ['行业','选择理由','首个标杆案例','A级行动场景'],
        [
          ['工程咨询','有招投标监控刚需，ROI直接可量化','广东企业筛出11亿线索','每日8:00前交付匹配准确率>90%的项目清单'],
          ['跨境电商','多语种多时区场景不可替代性强','中英双语双端宣传套图','一天内产出多语种宣传物料'],
          ['餐饮连锁','1人管多门店故事性极强','阿布拉江1家扩到6家','品牌营销+法务咨询+门店数据分析'],
        ]
      ),
      spacer(100),
      p('案例降维策略：每个行业选"最不像会用AI的客户"做标杆——连早餐店老板都能用，你为什么不能。'),
      spacer(),

      // ===== 三、指标体系 =====
      h1('三、指标体系'),
      divider(),

      h3('北极星：持续运营客户数'),
      p('定义：完成交付后连续60天有真实业务调用，且至少一个场景进入日常流程。六个月目标30家。'),
      p('为什么不看营收：行业数据85%企业完成POC但只有15%持续运营。标杆客户和案例产出是持续运营的前置条件——有了30家标杆，营收是自然结果。'),

      h4('辅助指标'),
      makeTable(
        ['指标','6个月目标','对应JD职责'],
        [
          ['可传播案例产出数','10个（覆盖3行业）','打造标杆案例…为核心内容支撑'],
          ['全链路转化率','到岗后定基线','拆解全链路转化漏斗，定位流失卡点'],
          ['FDE认证活跃数','100人','赋能销售、渠道与用户落地使用'],
        ]
      ),
      spacer(),

      // ===== 四、六个月执行路线 =====
      h1('四、六个月执行路线'),
      divider(),

      h3('第一阶段 · 1-30天 · 立信'),
      bullet('品类叙事定稿并全员对齐——销售、渠道、客服口径统一。核心一句话："不卖工具，卖用得起的AI落地服务"'),
      bullet('存量客户激活——安全侧客户经理发起500家企业筛选，主题是免费诊断而非产品介绍。发起"AI安全体检"，免费检测客户当前AI使用风险'),
      bullet('诊断工具包上线——20题标准问卷+报告模板+ROI测算器，渠道与FDE拿了就能用'),
      bullet('10家共创客户签署协议——给予优先功能、联合署名、创始价格，换取可公开的数据与故事使用权'),
      bullet('360内部改造故事首发——周鸿祎抖音发布"AI替我干了5个人的活"，讲述CEO超级助理真实经历'),

      h3('第二阶段 · 31-90天 · 验证'),
      bullet('交付包标准化——每行业沉淀"场景定义/配置模板/验收标准/常见问题/交付话术"'),
      bullet('3场老板行业闭门沙龙——客户现身说法20分钟→真实数据现场演示15分钟→FDE分桌诊断30分钟→当场预约'),
      bullet('案例生产线——每个标杆产出"短视频+图文+案例卡片+PPT页"四件套。禁止销售各自制作，统一标准化'),
      bullet('社媒矩阵启动——周鸿祎抖音"纳米Work实战周记"每周1条；小红书"用户故事"号每周2-3条；视频号深度演示'),

      h3('第三阶段 · 91-180天 · 放大'),
      bullet('渠道合伙人放大——交付包即准入教材，授权与认证等级绑定，未认证不得独立交付，防止交付质量崩塌反噬品牌'),
      bullet('城市服务站滚动深化——不追求同日开业的公关声量，按FDE产能开城；每城须有1名认证FDE与1个本地标杆'),
      bullet('内容矩阵规模化——同步做AEO/GEO，使AI搜索可检索到我方行业方法论'),
      bullet('老客扩展——对持续运营客户启动第二场景销售'),
      bullet('"100个行业的AI第一天"内容工厂启动——100个不同行业小企业主第一次用AI的完整记录'),
      spacer(),

      // ===== 五、内容与渠道 =====
      h1('五、内容与渠道'),
      divider(),

      h3('内容策略：从客户痛点场景出发，不从产品功能出发'),
      p('ToB不做大曝光，做信任飞轮。信任来自深度案例和行业口碑，不来自广告投放。先建痛点场景案例库，再让AI基于知识库生产内容。'),

      h3('渠道矩阵：每个平台打不同的牌'),

      h4('小红书："普通人用AI搞钱"的故事'),
      p('不做官方号发产品介绍，做"用户故事"号。选题："1人公司老板的一天：用AI管理6家门店""我用AI替代了3个外包""40岁餐饮老板学AI第1天到第30天"。找5-10个KOC免费试用，真实体验发笔记。'),

      h4('抖音：周鸿祎IP+场景化短视频'),
      p('固定栏目"纳米Work实战周记"每周1条。场景演示："用AI 5分钟筛出行业情报"。用户故事60-90秒，前5秒抛冲击观点。周鸿祎1500万粉丝，单条最高点赞53.1万。'),

      h4('视频号：B端决策者信息渠道'),
      p('360内部改造故事(30天改造360)、闭门沙龙精华剪辑、5-10分钟深度演示。视频结尾放企业微信二维码引导预约。B端决策者不怕长视频，怕的是内容不够硬。'),

      h4('公众号+知乎：深度案例+AEO占位'),
      p('公众号三层：信任层8篇(安全白皮书)+案例层10篇(深度案例)+选型层6篇(行业选型指南)。知乎找10-20个高赞问题专业回答。目标搜索"企业AI落地"时排前3。'),

      h4('行业垂直社区：精准触达'),
      p('工程咨询→工程造价论坛、招投标社群。跨境电商→雨果网、跨境眼。餐饮→餐饮老板内参、红餐网。不发产品介绍，发"行业AI解决方案"。'),

      calloutBox('💡 精妙点子：用纳米Work营销纳米Work','营销团队自己先用产品搭建营销智能体——自动生成社媒内容、分析用户反馈、生成案例报告。让营销过程本身成为产品能力的活案例。内容从客户痛点场景出发，先建场景案例库，再让AI基于知识库生产内容。'),
      spacer(),

      // ===== 六、内部协同 =====
      h1('六、内部协同机制'),
      divider(),
      p('360有4亿安全卫士+4亿浏览器+1200万纳米AI搜索+2万家政企客户。但跨部门协同不是"建议导流"就能实现的——没有利益分配就没有协同，这是大公司铁律。'),

      insightBox('◆ 协同四步法', [
        '利益绑——导流部门按Token消耗收入30%分成；安全销售交叉销售按15%计入安全团队业绩',
        '看板通——共享飞书/钉钉看板，所有部门看到实时数据（新增触达、注册、试用、深度验证、付费）',
        '周会盯——每周30分钟站会过5个数字。虚拟团队5人：产品+安全+搜索+品牌+渠道',
        '绩效挂——作战室成员季度绩效中20%与纳米Work整体指标挂钩。不改组织架构，只加考核维度',
      ]),
      spacer(100),
      p('参考：华为IPD的PDT跨部门团队，成员考核与产品商业结果强绑定；字节飞书×豆包团队整合，产品和销售端统一指挥。'),
      spacer(),

      // ===== 七、风险与应对 =====
      h1('七、风险与应对'),
      divider(),
      makeTable(
        ['风险','应对策略'],
        [
          ['公测用户转化率低\n参考Tome：2500万用户年收入仅300万','画像分层，从10万公测用户中筛选匹配企业版目标客群的几千家，集中FDE资源服务高价值种子'],
          ['FDE模式规模化天花板','设计轻量级自服务路径：预置A级场景智能体模板，技术强企业自主上手，FDE集中高价值客户'],
          ['竞品跟进安全/交付能力','把FDE网络做成不可复制的组织能力——代码能复制，人力组织能力不能'],
          ['10家共创客户交付效果不达预期','只宣传A级场景真实案例，B级标注"AI初稿+人工审核"，C级不做承诺。不overpromise'],
        ]
      ),
      spacer(100),
      h4('护城河时效'),
      makeTable(
        ['卖点','护城河深度','竞品追赶周期','原因'],
        [
          ['安全基因','深','2-3年','20年攻防积累，非代码可复制'],
          ['端到端交付','中','1-2年','任务引擎可被模仿，场景积累难复制'],
          ['FDE服务网络','最深','3-5年','人力组织能力，非技术可替代'],
        ]
      ),
      spacer(),

      // ===== 八、三个精妙点子 =====
      h1('八、三个精妙点子'),
      divider(),

      h4('1. AI数字员工入职'),
      p('从"买软件"变"招AI员工"。Token=试用期工资360付；FDE=入职培训师；500专家=AI员工简历库；城市合伙人=人才中介。企业对"招人"是刚需，对"买软件"有抵触。'),

      h4('2. 360 AI安全体检'),
      p('向2万家政企客户发"AI安全体检报告"，免费检测AI使用风险。从"你在卖AI工具"变成"你在帮我查安全漏洞"——心智完全不同。这不是推销，是安全服务延伸。'),

      h4('3. 100个行业的AI第一天'),
      p('100个不同行业的小企业主第一次用AI的完整记录。100条短视频+100篇图文=半年内容弹药库。"早餐店老板都能用，你是不是也该试试？"'),
      spacer(),

      // ===== 九、五个共识钩子 =====
      h1('九、五个共识钩子'),
      divider(),
      p('1. "ToB不做大曝光，做信任飞轮" — 信任来自深度案例和行业口碑，不来自广告投放'),
      p('2. "内容不从产品功能出发，从客户痛点场景出发" — 先建痛点场景案例库，再让AI基于知识库生产内容'),
      p('3. "利益绑、看板通、周会盯、绩效挂" — 四步法解决跨部门协同'),
      p('4. "小红书运营由懂ToB营销+懂AI+懂小红书的复合型人才主导" — 不是新媒体运营岗'),
      p('5. "用纳米Work营销纳米Work" — 营销团队自己先用产品，让营销过程成为产品能力的活案例'),
      spacer(400),

      // ===== 结语 =====
      h1('结语'),
      divider(),
      p('卖工具模式下，市场部的产出是线索；卖落地服务模式下，产出是可复制的交付资产与客户共识。'),
      p('纳米Work的营销不是把AI能力包装成功能列表卖出去，而是把360的安全基因、交付能力、FDE服务网络打包成"用得起的AI落地服务"——让企业老板从"不敢用AI"到"有人在帮我落地AI"。'),
      spacer(600),

      new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:600},children:[
        new ImageRun({data:logoBuffer,transformation:{width:30,height:30},type:'png'})
      ]}),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:100},children:[new TextRun({text:'纳米Work 企业版',bold:true,color:BLUE,size:24})]}),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:100},children:[new TextRun({text:'为企业AI化而生 · 从超级个体到超级组织',color:TEXT3,size:20})]}),
      new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'ToB不做大曝光，做信任飞轮',color:TEXT3,size:16,italics:true})]}),
    ],
  }],
});

Packer.toBuffer(doc).then(buffer=>{
  fs.writeFileSync('C:/Users/zhengcong/.qwenworkcn/workspace/msn0wsu702timscc/outputs/namiwork-marketing-plan-v2.docx',buffer);
  console.log('Done');
});

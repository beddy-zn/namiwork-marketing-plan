const docx = require('docx');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType, BorderStyle, PageBreak, Footer, Header, PageNumber, ImageRun } = docx;
const fs = require('fs');
const path = require('path');

const logoBuffer = fs.readFileSync(path.join(__dirname, 'hex-logo.png'));
const BLUE='1B3A6B',BLUE_DARK='0F2847',BLUE_LIGHT='E8F0FE',GOLD='D4A020';
const TEXT='1A1A2E',TEXT2='4A5568',TEXT3='8A94A6';

function p(t,o={}){const{bold,color,size,italic,align,spacing}=o;return new Paragraph({spacing:{after:spacing||200,line:360},alignment:align||AlignmentType.LEFT,children:[new TextRun({text:t,bold:bold||false,color:color||TEXT2,size:size||22,italics:italic||false})]})}
function h1(t){return new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{before:400,after:200},children:[new TextRun({text:t,bold:true,color:BLUE_DARK,size:36})]})}
function h2(t){return new Paragraph({heading:HeadingLevel.HEADING_2,spacing:{before:360,after:160},children:[new TextRun({text:t,bold:true,color:BLUE,size:28})]})}
function h3(t){return new Paragraph({heading:HeadingLevel.HEADING_3,spacing:{before:280,after:120},children:[new TextRun({text:t,bold:true,color:TEXT,size:26})]})}
function h4(t){return new Paragraph({heading:HeadingLevel.HEADING_4,spacing:{before:200,after:80},children:[new TextRun({text:t,bold:true,color:TEXT,size:22})]})}
function divider(){return new Paragraph({spacing:{before:100,after:200},border:{bottom:{color:BLUE,size:6,style:BorderStyle.SINGLE,space:1}},children:[]})}
function spacer(a=200){return new Paragraph({spacing:{after:a},children:[]})}
function bullet(t){return new Paragraph({spacing:{after:80,line:340},bullet:{level:0},children:[new TextRun({text:t,color:TEXT2,size:21})]})}

function insightBox(label,paras){
  return new Table({width:{size:100,type:WidthType.PERCENTAGE},columnWidths:[9000],rows:[
    new TableRow({children:[new TableCell({
      width:{size:100,type:WidthType.PERCENTAGE},shading:{type:ShadingType.CLEAR,fill:BLUE_LIGHT},
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
      width:{size:100,type:WidthType.PERCENTAGE},shading:{type:ShadingType.CLEAR,fill:'FFF9E6'},
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
      new Paragraph({spacing:{after:300},alignment:AlignmentType.CENTER,children:[new TextRun({text:'MARKETING STRATEGY · 2026',color:BLUE,size:20,bold:true})]}),
      new Paragraph({spacing:{after:300},alignment:AlignmentType.CENTER,children:[new TextRun({text:'六个月GTM营销方案',color:TEXT,size:48,bold:true})]}),
      new Paragraph({spacing:{after:400},alignment:AlignmentType.CENTER,children:[new TextRun({text:'不在办公入口与BAT争夺存量，以「企业AI落地服务」重新定义品类\n用安全资产做信任杠杆，用FDE网络做交付壁垒',color:TEXT3,size:18,italics:true})]}),
      new Paragraph({spacing:{after:100},alignment:AlignmentType.CENTER,children:[new TextRun({text:'2026年8月',color:TEXT3,size:18})]}),
      new Paragraph({children:[new PageBreak()]}),

      // ===== 00 核心战略判断 =====
      h1('00  核心战略判断'),
      divider(),

      h3('一、判断'),
      p('纳米Work企业版不应该在「AI办公智能体」赛道正面竞争。钉钉、飞书、企微的护城河不是AI能力，是组织已经在里面办公——通讯录、审批流、文件、历史沉淀都在那儿。要求客户先离开一个每天都在用的系统，这件事营销解决不了。'),

      h3('二、打法'),
      p('把赛道重新定义为「企业AI落地服务」，对手换成收费几十万的咨询公司与SI集成商。在这条赛道上，360的三项资产——安全信任、端到端交付、FDE服务网络——是竞品短期内难以补齐的。'),

      h3('三、六个月目标'),
      bullet('30家持续运营客户（完成交付后连续60天有真实业务调用，且至少一个场景进入日常流程）'),
      bullet('9个可复制交付包（3行业 × 3场景）'),
      bullet('10个可传播案例'),

      h3('四、核心风险'),
      p('交付跟不上。FDE产能是这套打法的关键瓶颈——营销跑在产能前面等于签了没人交付，口碑反噬比签不到单代价更大。整个方案的设计逻辑是：先拿标杆验证交付，再标准化复制，最后才放大。'),
      spacer(),

      // ===== 01 市场判断与品类定义 =====
      h1('01  市场判断与品类定义'),
      divider(),

      h3('1.1  三家竞品的模式——我们不打'),
      p('钉钉→千问办公、飞书→Aily、企业微信→WorkBuddy，三家BAT都是"办公生态+嵌入AI智能体"模式。它们的护城河不是AI，是组织已经在里面办公。正面进攻等于要求客户换一套每天都在用的系统。'),
      p('其他竞品：WPS灵犀、TRAE Work、Microsoft 365 Copilot、ChatGPT、腾讯元宝、豆包、Kimi、Notion AI。这些产品要么绑定自有生态，要么是通用AI助手，在企业落地服务层面缺乏深度。'),

      h3('1.2  品类重新定义'),
      p('从「卖工具」切换到「卖用得起的AI落地服务」。在工具赛道上客户会逐条比功能，在落地服务赛道上客户比的是「谁能把这件事做成」。叙事框架：超级个体→超级组织。'),

      h3('1.3  三项护城河，以及各自的真实边界'),

      h4('安全 = 信任底座'),
      p('不是功能卖点，而是准入门槛。360累计捕获60个境外APT组织，3800+安全专家团队，覆盖2万家政企客户。市场上63.6%的AI供应商未在法律文件中披露第三方数据转发，企业用AI最大的顾虑就是数据安全。360的安全能力不是后贴的，是20年攻防积累的原生能力。'),
      p('边界：安全是门槛不是卖点——客户因为信任你才愿意试，但最终留不留得住看交付效果。'),

      h4('交付 = 服务承诺'),
      p('竞品交付"你用AI做出来的"，纳米Work交付"AI替你做出来的"。纳米1314任务引擎：用户输入目标→AI拆解任务→调度多专家→云端7×24执行→交付成品。'),

      h4('交付场景分级'),
      makeTable(
        ['级别','定义','代表场景','交付承诺'],
        [
          ['A级','可全自动交付','招投标监控、多语种内容生成、合规台账','7×24无人值守'],
          ['B级','AI交付+人工审核','合同审查、营销方案生成、客户分析','AI完成80%，人工判断'],
          ['C级','AI辅助+人主导','战略规划、创意决策','提供信息支持，不做交付承诺'],
        ]
      ),
      spacer(100),
      p('边界：A级场景才做"端到端交付"的承诺，B级标注"AI初稿+人工审核"，C级不承诺。不overpromise是信任的底线。'),

      h4('FDE = 核心产能'),
      p('依托360白帽子群体→FDE转型+城市合伙人=网格式AI落地服务网络。FDE三级认证：训练工匠→AI业务专家→AI咨询架构师。白帽子转型需要分三阶段：学AI产品操作（2周）→跟单学业务理解（4周）→独立交付学项目管理（8周）。'),
      p('边界：FDE是人力密集型交付，规模化天花板明显。方案设计轻量级自服务路径做分流。'),
      spacer(),

      // ===== 02 目标客户与优先级 =====
      h1('02  目标客户与优先级'),
      divider(),

      h3('2.1  把安全客户变成AI客户'),
      p('首先从存量安全客户库转化——已付费、已信任、已有采购关系与对接人。按三条件筛出首批300家：'),
      bullet('员工50-500人（中小企业）'),
      bullet('近12个月有续费记录'),
      bullet('属于三个Beachhead行业'),
      p('由安全侧客户经理带FDE二次拜访，不做产品推销，做免费「AI场景诊断」。从"你在卖AI工具"变成"你在帮我看AI怎么用安全"——心智完全不同。'),

      h3('2.2  三个Beachhead行业建议'),
      makeTable(
        ['行业','选择理由','首个标杆案例','A级交付场景'],
        [
          ['工程咨询','有招投标监控刚需，ROI可量化','广东企业筛出11亿线索','每日8:00前交付匹配准确率>90%的项目清单'],
          ['跨境电商','多语种多时区场景不可替代','中英双语双端宣传套图','一天内产出多语种宣传物料'],
          ['餐饮连锁','1人管多门店故事性强','阿布拉江1家扩到6家','品牌营销+门店数据分析+法务咨询'],
        ]
      ),
      spacer(100),
      p('案例降维策略：每个行业选"最不像会用AI的客户"做标杆——连早餐店老板都能用起来，说明这件事不需要技术门槛。'),
      spacer(),

      // ===== 03 指标体系 =====
      h1('03  指标体系'),
      divider(),

      h3('3.1  北极星指标'),
      p('持续运营客户数。定义：完成交付后连续60天有真实业务调用，且至少一个场景进入日常流程。目标30家。'),
      p('为什么不看营收：行业数据显示85%企业能完成POC但只有15%能持续运营。前6个月的核心是进入那15%，有了持续运营的客户，营收是自然结果。'),

      h3('3.2  辅助指标'),
      makeTable(
        ['指标','6个月目标','说明'],
        [
          ['可复制交付包','9个（3行业×3场景）','每包含场景定义/配置模板/验收标准/交付话术'],
          ['可传播案例','10个','覆盖3行业，每个案例产出短视频+图文+数据页三件套'],
          ['FDE认证活跃数','100人','保障交付产能的关键供给指标'],
          ['全链路转化率','到岗后定基线','从触达到深度试用的各环节转化'],
        ]
      ),
      spacer(),

      // ===== 04 六个月执行路线 =====
      h1('04  六个月执行路线'),
      divider(),
      p('共三个阶段，每阶段只认一个关键结果，并预设退出条件。'),

      h3('4.1  第一阶段 · 1-30天 · 立信'),
      p('关键结果：10家共创客户签署协议。退出条件：如果前30天签不到5家，说明品类叙事或客户筛选标准需要调整，不进入下一阶段。'),
      bullet('品类叙事定稿并全员对齐——销售、渠道、客服口径统一。核心一句话："不卖工具，卖用得起的AI落地服务"'),
      bullet('存量客户激活——安全侧客户经理按三条件筛出500家可触达池，主题是免费AI场景诊断而非产品介绍'),
      bullet('诊断工具包上线——20题标准问卷+报告模板+ROI测算器，渠道与FDE拿了就能用'),
      bullet('10家共创客户签署协议——给予优先功能、联合署名、创始价格，换取可公开的数据与故事使用权'),
      bullet('360内部改造故事首发——周鸿祎抖音发布"用AI改造360的30天"，讲述CEO超级助理的真实经历'),

      h3('4.2  第二阶段 · 31-90天 · 验证'),
      p('关键结果：3场闭门沙龙+9个交付包标准化。退出条件：如果交付包无法标准化（每个客户都要从头定制），说明产品成熟度不够，暂停放大。'),
      bullet('交付包标准化——每行业沉淀「场景定义/配置模板/验收标准/常见问题/交付话术」'),
      bullet('3场老板级行业闭门沙龙——客户现身说法20分钟+真实数据现场演示15分钟+FDE分桌诊断30分钟+当场预约'),
      bullet('案例生产线——每个标杆产出短视频+图文+数据页三件套，统一标准化，禁止销售各自制作'),
      bullet('社媒矩阵启动——周鸿祎抖音"纳米Work实战周记"每周1条；小红书"用户故事"号每周2-3条；视频号深度演示'),

      h3('4.3  第三阶段 · 91-180天 · 放大'),
      p('关键结果：渠道合伙人签约5城+老客第二场景启动。退出条件：如果未认证FDE开始独立交付且出现质量事故，暂停该城市运营。'),
      bullet('渠道合伙人放大——授权与认证等级绑定，未认证不得独立交付，防止交付质量崩塌反噬品牌'),
      bullet('城市服务站滚动深化——不追求同日开业的公关声量，按FDE产能开城；每城须有1名认证FDE与1个本地标杆'),
      bullet('内容矩阵规模化——同步做AEO/GEO，让AI搜索引擎能检索到我方行业方法论'),
      bullet('老客扩展——对持续运营客户启动第二场景销售，这是六个月后收入曲线的起点'),
      spacer(),

      // ===== 05 内容与渠道 =====
      h1('05  内容与渠道'),
      divider(),

      h3('5.1  内容策略'),
      p('ToB不做大曝光，做信任飞轮。信任来自深度案例和行业口碑。内容不从产品功能出发，从客户痛点场景出发——先建痛点场景案例库，再让AI基于知识库生产内容。'),
      p('每个标杆案例的标准化产出（三件套）：3分钟短视频（抖音/视频号）+ 深度图文（公众号/知乎）+ 数据页（销售/渠道使用）。统一用Before/After结构：用AI前花多少人多少时间，用后变成多少。'),

      h3('5.2  渠道矩阵'),
      makeTable(
        ['平台','定位','核心内容方向','节奏'],
        [
          ['抖音','周鸿祎IP+场景化演示','"纳米Work实战周记"固定栏目+用户故事60秒','每周1条'],
          ['小红书','普通人用AI搞钱的故事','"1人公司老板的一天""我用AI替代了3个外包"',  '每周2-3条'],
          ['视频号','B端决策者渠道','360内部改造故事+沙龙精华+5-10分钟深度演示','每周1条'],
          ['公众号','深度案例+行业解决方案','三层：信任层8篇+案例层10篇+选型层6篇','每周1篇'],
          ['知乎','AEO搜索占位','找10-20个"企业AI"高赞问题，专业身份回答','每周2篇'],
          ['行业垂直社区','精准触达','工程造价论坛/雨果网/餐饮老板内参发行业解决方案','每月2篇'],
        ]
      ),
      spacer(100),
      calloutBox('用纳米Work营销纳米Work','营销团队自己先用产品搭建营销智能体——自动生成社媒内容、分析用户反馈、生成案例报告。让营销过程本身成为产品能力的证明。'),
      spacer(),

      // ===== 06 风险与竞争响应 =====
      h1('06  风险与竞争响应'),
      divider(),
      makeTable(
        ['风险','应对策略'],
        [
          ['FDE产能跟不上\n营销跑在交付前面','整个方案按"先验证再放大"设计。第一阶段只签10家共创，不追求量。轻量级自服务路径做分流'],
          ['交付效果不达预期\n客户觉得"AI不靠谱"','只宣传A级场景真实案例，B级标注"AI初稿+人工审核"。不overpromise是底线'],
          ['公测用户转化率低\n10万用户多为薅羊毛','画像分层，从10万中筛出匹配企业版客群的几千家，集中FDE资源服务高价值种子'],
          ['竞品跟进安全/交付能力','把FDE网络做成组织能力——代码能复制，人力组织能力不能。安全基因20年积累也是壁垒'],
        ]
      ),
      spacer(100),
      h4('护城河时效'),
      makeTable(
        ['卖点','护城河深度','竞品追赶周期','原因'],
        [
          ['安全基因','深','2-3年','20年攻防积累，非代码可复制'],
          ['端到端交付','中','1-2年','任务引擎可被模仿，场景积累需要时间'],
          ['FDE服务网络','最深','3-5年','人力组织能力，非技术可替代'],
        ]
      ),
      spacer(),

      // ===== 07 组织与资源 =====
      h1('07  组织与资源'),
      divider(),

      h3('7.1  市场部角色的变化'),
      p('卖工具模式下，市场部的产出是线索；卖落地服务模式下，产出是可复制的交付资产与客户共识。这意味着市场部的考核从MQL数量转向交付包与案例的产出，且必须与FDE团队共背北极星指标。'),

      h3('7.2  跨部门协同——四步法'),
      p('360有4亿安全卫士月活、4亿浏览器月活、1200万纳米AI搜索月活、2万家政企客户。但跨部门协同不是"建议导流"就能实现的——没有利益分配就没有协同。'),
      insightBox('利益绑、看板通、周会盯、绩效挂', [
        '利益绑——导流部门按Token消耗收入30%分成；安全销售交叉销售按15%计入安全团队业绩',
        '看板通——共享飞书/钉钉看板，所有部门看到纳米Work实时数据',
        '周会盯——每周30分钟站会过5个数字。虚拟团队5人：产品+安全+搜索+品牌+渠道',
        '绩效挂——作战室成员季度绩效中20%与纳米Work整体指标挂钩。不改组织架构，只加考核维度',
      ]),
      spacer(100),
      p('参考：华为IPD的PDT跨部门团队，成员考核与产品商业结果强绑定；字节飞书×豆包团队整合，产品和销售端统一指挥。'),

      h3('7.3  人力配置建议'),
      p('小红书运营不是交给新媒体运营做，而是由懂ToB营销+懂AI实操+懂小红书内容的复合型人才主导。ToB产品的社媒内容不是追热点，是建信任——需要既理解企业客户痛点、又懂AI产品能力、还懂内容运营的人来做。'),
      spacer(),

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
    ],
  }],
});

Packer.toBuffer(doc).then(buffer=>{
  fs.writeFileSync('C:/Users/zhengcong/.qwenworkcn/workspace/msn0wsu702timscc/outputs/namiwork-marketing-plan-v4.docx',buffer);
  console.log('Done');
});

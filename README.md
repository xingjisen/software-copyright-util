# 软著申请资料生成 Skill

> 从真实项目源码出发，一步步生成可提交的**软件著作权申请材料**：申请表信息、程序鉴别材料（源代码前后 30 页）、文档鉴别材料（操作手册），以及配套的软件界面截图。

仓库地址：<https://github.com/xingjisen/software-copyright-util>

面向中国版权保护中心的计算机软件著作权登记场景。整个流程把"读代码 → 理解业务 → 写材料 → 排版成 Word → 实测页数"串成一条可追溯的流水线，每一步都要用户确认后才继续，不会绕过人工判断自行编造内容。

---

## 一、它解决什么问题

软著申请的坑不在填表，而在于材料的"一致性"和"形式合规"：

- 申请表里写的功能，操作手册里必须真的有对应页面；
- 操作手册的每个页面，源码里必须真的有对应实现；
- 代码材料要求前 30 页 + 后 30 页、每页不少于 50 行，且页眉软件名称版本号必须与申请表一致；
- 操作手册一般不少于 40 页，且要写得像真实软件随附的说明书，不能像技术文档。

人工做这些事，最耗时的是反复对齐和排版核对。这个 skill 把这些规则固化进脚本，自动抽取代码、自动排版、自动实测页数，把模型的精力留在"业务理解"和"文字质量"上。

## 二、两种使用模式

| 模式 | 适用场景 | 流程 |
|---|---|---|
| **已有源码**（常规） | 手上有一个真实项目，要申请软著 | 分析项目 → 业务理解 → 代码抽取 → 申请表 → 操作手册 → 截图 → Word |
| **从零起步**（bootstrap） | 只有软件名称和想法，还没有代码 | 需求访谈 → 按页面清单生成静态 HTML 原型 → 用户逐页确认 → 按原型生成配套源码 → 再走常规流程 |

从零起步模式生成的源码属于**用户委托创作**，与原型、需求一一对应，不套用任何现成项目代码。

## 三、工作流程

```
1. 环境检查        检查 Python / python-docx / .NET SDK，给出降级方案
2. 定位项目        确认源码目录
2A. 从零起步       （仅 bootstrap 模式）需求访谈 → 原型 → 源码
3. 分析项目        扫描文件结构、语言构成、入口与模块
4. 形成业务理解    行业、目标用户、核心功能、申请口径
5. 确认申请表字段  软件全称、版本、著作权人、开发环境……
6. 确认代码选择    从候选清单里挑"最能体现业务"的文件，凑成整页
7. 生成草稿        Markdown 版申请表信息、操作手册、代码材料
8. 获取截图        自动截图 / 用户提供 / 先跳过
9. 确认草稿        用户逐份审阅
10. 生成正式文档   Word + TXT
11. 三轮验证       页数实测、页眉一致性、内容自检
```

### 人工门禁（关键设计）

每个关键节点都会停下等用户确认，并把确认结果落盘为门禁 JSON，中断后可以从断点续跑：

| 门禁 | 确认内容 |
|---|---|
| `requirements` | 需求说明的功能清单与业务逻辑 |
| `prototype` | 原型页面内容与业务一致 |
| `codegen` | 技术栈与源码目录 |
| `business` | 行业、目标用户、业务功能、申请口径 |
| `code-selection` | 抽取哪些代码文件及理由 |
| `application-fields` | 申请表全部字段（含著作权人） |
| `screenshot-method` | 截图获取方式 |
| `markdown` | 全部草稿内容 |

## 四、产出物

生成在用户项目的 `软件著作权申请资料/` 目录下：

| 文件 | 说明 |
|---|---|
| `正式资料/申请表信息.txt` | 全部申请表字段，逐项对照在线申请系统填写 |
| `正式资料/<软件名>_操作手册.docx` | 操作手册，含目录、页面截图、操作步骤、状态速查表（≥40 页） |
| `正式资料/<软件名>-代码(前后30页).docx` | 程序鉴别材料，前 30 页 + 后 30 页，每页 50 行，页眉带软件名与版本 |
| `草稿/需求说明.md` | bootstrap 模式下的需求说明 |
| `草稿/业务理解.md` | 业务理解与申请口径 |
| `草稿/操作手册.md`、`申请表信息.md`、`代码-前30页+后30页.md` | 可编辑的 Markdown 草稿 |
| `原型/*.html` | bootstrap 模式下的静态原型页面（可双击用浏览器打开） |
| `截图/*.png` | 各页面截图，已插入手册对应章节 |

## 五、安装

从仓库克隆到技能目录（用户级，所有项目可用）：

```bash
cd ~/.workbuddy/skills
git clone https://github.com/xingjisen/software-copyright-util.git
```

也可以手动把整个目录放到下列位置：

```bash
# 用户级（所有项目可用，推荐）
~/.workbuddy/skills/software-copyright-util/

# 或项目级（仅当前项目）
<项目>/.workbuddy/skills/software-copyright-util/
```

Windows 下即 `C:\Users\<用户名>\.workbuddy\skills\software-copyright-util\`。

## 六、环境要求

| 依赖 | 是否必须 | 说明 |
|---|---|---|
| Python 3.10+ | 必须 | 运行全部流程脚本 |
| `python-docx` | 必须 | `pip install python-docx`，生成 Word |
| .NET SDK 8.0+ | 可选 | 完整 OpenXML 校验与预览，跑 `vendor/docx-toolkit/scripts/setup.sh` |
| Chrome / Edge | 截图时需要 | 原型页面渲染与截图 |

缺 .NET 时走兜底路径：仍能生成 Markdown、TXT 和基础 DOCX，报告中会注明当前使用兜底。

## 七、怎么用

直接用自然语言提需求即可，skill 会自动接管：

```
帮我给这个项目申请软著            # 已有源码
读取当前目录的项目，生成软著申请资料
我只有软件名，想先做一套软著材料   # 从零起步
```

过程中它会一次只问一个问题（比如著作权人是谁、技术栈选哪个），按提示回答即可。**材料内容全部来自你的确认**，不会替你决定著作权人、发表状态这类法律事实。

## 八、目录结构

```
software-copyright-util/
├── SKILL.md                  技能主文档：流程、门禁、脚本用法、规则
├── README.md                 本文件：面向人的介绍
├── .gitignore                排除缓存与 .NET 构建产物
├── agents/
│   └── openai.yaml           技能展示名与默认提示词
├── references/               细则规则（主文档按需引用）
│   ├── application_fields.md            申请表字段清单
│   ├── business_understanding_rules.md  业务理解证据规则
│   ├── code_selection_rules.md          代码抽取选择规则
│   ├── copyright_material_rules.md      软著材料法定要求
│   └── manual_structure.md              操作手册骨架
├── scripts/                  流程脚本
│   ├── check_environment.py        环境检查
│   ├── analyze_project.py          项目分析
│   ├── confirm_stage.py            门禁记录与断点恢复
│   ├── generate_business_context.py 业务理解
│   ├── propose_code_selection.py   代码候选清单
│   ├── extract_code_material.py    代码材料抽取（前后 30 页）
│   ├── generate_application_info.py 申请表信息
│   ├── generate_manual_draft.py     操作手册草稿
│   ├── capture_screenshots.py       截图整理
│   ├── build_docx_from_md.py        Markdown → Word
│   └── measure_docx_pages.py        Word 页数实测
└── vendor/
    └── docx-toolkit/         内置 DOCX 工具包（OpenXML，MIT 许可）
```

## 九、设计原则

- **代码不编造**：程序鉴别材料只从用户确认的完整源码文件里抽取，禁止拼接或生成假代码。唯一例外是从零起步模式，此时源码本身就是按用户需求和原型委托创作的。
- **页数不靠估算**：所有页数结论都以 `measure_docx_pages.py` 实测为准，估算值只用于中间校验。
- **内容不空转**：操作手册按真实页面逐章展开，写清模块用途、操作过程与结果反馈，避免"点击按钮完成操作"这类空话。
- **每次停下确认**：宁可多问一句，也不替用户拍板。

## 十、版本记录

| 版本 | 变更 |
|---|---|
| v1.0 | 首个发布版本：需求访谈与从零起步模式、原型页面、源码生成、业务理解、申请表信息、代码材料（前后 30 页）、操作手册、截图、Word/TXT 输出全流程打通；原型视觉规范（颜色只做点缀不做容器、白底卡片、布局范式、信息密度下限、交互联动范围、原型自检清单）；抽取脚本对被跳过文件显式告警，Word 生成前自动清理手册尾部提示块 |

## 十一、来源与许可

- 本项目仓库：[xingjisen/software-copyright-util](https://github.com/xingjisen/software-copyright-util)，作者 xingjisen。
- 本 skill 基于 [Fokkyp/SoftwareCopyright-Skill](https://github.com/Fokkyp/SoftwareCopyright-Skill) 演进，感谢原作者的开源工作。
- `vendor/docx-toolkit/` 为第三方组件，采用 MIT 许可，其 `LICENSE` 文件随目录一同保留。
- 使用前请自行确认所在地区软著登记的最新要求；本工具产出的是材料草稿与排版成品，**法律事实（著作权人、开发日期、发表状态等）须由申请人自行确认**。

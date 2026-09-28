# CDR X4 增强工具包（CDRX4Toolkit）

**v1.2.0** · MIT · 适用于 **CorelDRAW X4**（需带 VBA 组件）

把「CorelDRAW X4 精简增强版」里那批插件的常用功能，用 VBA 宏（GMS）**独立重写**了一遍，打包成**双击即用**的单个 `.vbs`，可以自由分发给任意普通 X4。

装好后 CDR 里多出一条「**增强工具**」工具栏，9 个中文按钮，**打开就在，不用手动配置**。

- 仓库：<https://github.com/Ri1035/cdrx4-toolkit>
- 全部源码开放，可自行修改、打包、分发
- 变更记录：[CHANGELOG.md](CHANGELOG.md) ｜ 修复过程与技术实录：[PLAN.md](PLAN.md)

---

## 一、快速开始

### 安装

**双击 `dist\安装CDRX4增强工具.vbs`**（从 GitHub 用，就只下载这一个文件）。

1. 装之前**先关掉 CorelDRAW**
2. 双击安装器，它会弹窗告诉你结果
3. 重新打开 CorelDRAW → 「增强工具」工具栏出现，9 个按钮直接可用

安装器做两件事，**全程不碰运行中的 CorelDRAW**：

1. 把内嵌的 `CDRX4Toolkit.gms` 写进
   `%APPDATA%\Corel\CorelDRAW Graphics Suite X4\User Draw\GMS\`
2. 把「增强工具」工具栏的 XML 写进 CorelDRAW 工作区文件
   `...\User Workspace\<应用>\<工作区>\DRAWUIConfig.xml`（改前先备份成 `.cdrx4bak`）

它只认 **X4 系列**目录，不会污染 2022 等其它版本。

> 第 2 步不是「绕过」CorelDRAW，而是**走 CorelDRAW 自己的路**：这条工具栏标记与你在
> 「工具 > 自定义 > 命令 > 宏」里把宏拖到工具栏上时 CorelDRAW 写的完全一样，
> 所以按钮照常执行宏，而安装器本身不可能让 CDR 崩。原因见 [技术要点](#六技术要点改代码前必读)。

### 卸载

**双击 `dist\卸载CDRX4增强工具.vbs`**。它做两件事：

1. 删掉所有 X4 目录下的 `CDRX4Toolkit.gms`（含被手工改名成 `.bak` 的）
2. 把「增强工具」工具栏的标记从工作区文件里清掉（改前备份成 `.uninstallbak`）

**跑之前必须完全关闭 CorelDRAW** —— CDR 退出时会用自己的设置覆盖工作区文件，边运行边改会被冲掉。
脚本检测到 CDR 还开着会**直接拒绝执行**并告诉你怎么办（0 线程的崩溃僵尸不算「在运行」）。

手工卸载等价于上面两步：删掉那个 GMS → 重启 CDR → 右键「增强工具」工具栏 → 删除。

### 清掉旧版（v1.0.2 及更早）

**用同一个卸载器就能清干净。** 它不认版本号，只按两个稳定特征定位：
`itemData` 的宏名里含 `CDRX4Toolkit`、工具栏名叫「增强工具」——
所以旧版（包括用 `CommandBars` API 建的、guid 随机的那条工具栏）留下的东西一样删得掉。

清完再双击新版安装器即可。

> 装了 v1.0.2 的设备上 CDR 一启动就卡死，根因就在旧版的启动钩子。详见 [技术要点](#六技术要点改代码前必读)。

### 前提：必须有 VBA 组件

GMS 依赖 VBA。完整版 X4 自带；如果按 `Alt+F11` 打不开 VBA 编辑器，说明 VBA 被裁掉了，
需要先补装 VBA 组件，否则本包无法运行。

---

## 二、功能

工具栏上的按钮顺序即下表顺序（前三个是高频功能，排在前面）：

| # | 按钮 | 用法 | 说明 |
|---|---|---|---|
| 1 | **全部转曲** | 直接点 | 遍历所有页面的对象转曲。**弹窗两步**：① 范围（全部 / 仅文字 / 仅图形）② 是否深入 PowerClip 容器 |
| 2 | **转CMYK** | 直接点 | 把 RGB 等填充 / 轮廓转 CMYK；可选纯黑叠印 |
| 3 | **颜色替换** | 先选中 **2 个**对象 | 第 1 个 = 原色，第 2 个 = 新色；全文档替换填充与轮廓 |
| 4 | **标准矩形** | 先选中矩形 | 把圆角 / 扇形角 / 倒棱角重置为直角 |
| 5 | **对象适合路径** | 选中 **1 条曲线 + 2 个对象** | 把对象副本沿路径节点均匀排布（见下） |
| 6 | **JPG批量导出** | 直接点 | 按页导出 JPG，可设文件夹与 dpi |
| 7 | **插入页码** | 直接点 | 每页底部居中插入「第 X 页 / 共 Y 页」 |
| 8 | **日历创建** | 直接点 | 输入 `2026-09`，生成该月文字日历 |
| 9 | **印章制作** | 直接点 | 外圆 + 环形文字 + 五角星 + 下部文字 |

### 与原插件的差异

- **转CMYK**：原版对话框还带 `分辨率 / 反锯齿 / 透明背景 / ICC 配置文件` 等选项，那些是**导出参数**；
  本版只做色彩模型转换 + 叠印。需要「导出型转 CMYK」可以再加。
- **对象适合路径**：X4 的 `CreateBlend` 对外部调用一律报类型不匹配（已实测），所以本版不走混合，
  改为**按路径节点 / 线段均匀插值排布对象副本**。结果是散开的独立对象，不是混合组。
- **防伪花纹 / 防伪轮廓**：未实现（那是专门的防伪底纹引擎，也不建议复刻）。

### 建议人工验收的点

9 个功能已在干净的 X4 会话里跑过自动化冒烟，全部 `err=0`；但自动化只验证「不报错、处理了对象」，
**视觉效果**请各点一遍确认：

1. 圆 / 五角星 / 日历文字的**坐标与对齐**
2. **JPG批量导出**的翻页与文件命名
3. **对象适合路径**的间距是否均匀、朝向是否合适

---

## 三、常见问题

**装完 CDR 一启动就卡死 / 无响应，连关都关不掉**
那是 v1.0.2 的 bug，不是这台机器坏了。按顺序做：
① 用任务管理器结束 `CorelDRW.exe`；② 结束不掉（进程变成 0 线程的僵尸）就**重启一次电脑**；
③ 双击 `卸载CDRX4增强工具.vbs` 清掉旧版；④ 再装 v1.2.0。
v1.0.3 起已经没有启动钩子，不会再出现这个问题。

**重启 CDR 后没看到「增强工具」工具栏**
按可能性排查：① 安装时 CDR 没关 —— 关掉重装；② 工作区写入失败（权限不足），
安装器弹窗会明确提示，此时可在 CDR 里手动跑一次宏 `CDRX4Toolkit > M_Install > InstallToolbar`，
或走「工具 > 自定义 > 命令 > 宏」把宏拖到工具栏上；③ 工具栏被拖到屏幕外或隐藏了，看「工具 > 自定义」。

**工具栏在，但按钮点不动 / 报「找不到命令」**
说明 GMS 没被加载。按 `Alt+F11` 看有没有 `CDRX4Toolkit` 工程 —— 没有就是 VBA 组件缺失，
或机器上已有同名的 `CDRX4Toolkit` 工程冲突（打包分发前先清掉重名工程）。

**为什么 `CorelDRW.exe` 杀不掉**
那是崩溃残留的 **0 线程僵尸**，`taskkill /F` 会报 Access denied，**只能重启系统**才消失。
它不影响新启动的 CDR，也不影响安装器 / 卸载器（两者都用 `ThreadCount > 0` 判断「真的在运行」）。

**能装到 CorelDRAW 2022 吗**
不能，本包只认 X4 系列目录。2022 自带 `ConvertAllToCurves.gms`，也有更成熟的开源方案（见下）。

**弹窗中文变成 `??????` / 双击安装器报「无效字符」（800A0408）**
都是旧版 bug，v1.0.1 起已修复：文案改为 base64 内嵌、运行时解码，被执行的 `.vbs` 一律纯 ASCII。

---

## 四、目录结构

```
cdrx4-toolkit/
├─ README.md                  ← 本文件
├─ CHANGELOG.md               ← 版本变更记录
├─ LICENSE                    ← MIT 开源协议
├─ .gitignore
├─ dist/                      ← 分发就发这个文件夹里的 .vbs
│  ├─ 安装CDRX4增强工具.vbs    ← 双击安装（内嵌插件本体 + 工具栏标记，单文件可用）
│  ├─ 卸载CDRX4增强工具.vbs    ← 双击卸载（删插件本体 + 清工作区里的工具栏）
│  └─ CDRX4Toolkit.gms        ← 插件本体（build_gms.vbs 生成，二进制未入库）
├─ src/                       ← 9 个功能模块 + 安装模块 + 自检模块的源码（UTF-8）
│  ├─ M_Util.bas              ← 公共函数（单位换算 / 颜色 / 定位 / 文本）
│  ├─ M_Install.bas           ← 工具栏安装 / 卸载（按钮中文名与顺序都在这里）
│  ├─ M_Curves.bas            ← 全部转曲
│  ├─ M_CMYK.bas              ← 转CMYK
│  ├─ M_Color.bas             ← 颜色替换
│  ├─ M_Rect.bas              ← 标准矩形
│  ├─ M_FitPath.bas           ← 对象适合路径
│  ├─ M_JPG.bas               ← JPG批量导出
│  ├─ M_PageNo.bas            ← 插入页码
│  ├─ M_Calendar.bas          ← 日历创建
│  ├─ M_Seal.bas              ← 印章制作
│  ├─ M_Test.bas              ← 自检：跑 9 个功能并写日志，供冒烟测试判成败
│  └─ _startup.txt            ← 文档模块的启动代码；默认 `--none`（= 不写任何启动代码）
├─ logs/                      ← 验证日志（真实启动测试 / 冒烟 / 构建）
├─ tools/                     ← 开发 / 验证脚本（不参与运行，可删）
│  ├─ smoke.vbs               ← 冒烟测试：重启 X4，验工程 / 工具栏 / 9 个功能
│  ├─ real_launch_test.ps1    ← 真实启动测试：手动启动 CorelDRW.exe，看进程健康度 + 崩溃事件
│  ├─ workspace_inject.ps1    ← 往工作区注入工具栏标记（原型，安装器已内置同样逻辑）
│  ├─ workspace_strip.ps1     ← 把工具栏标记从工作区剥掉，用来模拟「干净机器」
│  └─ probe*.vbs              ← X4 对象模型探针（当初用来核对 API，已归档）
├─ build_gms.vbs              ← 用 src/ 里的源码生成 GMS（需本机装 X4）
├─ workspace_markup.txt       ← 工具栏的三段 XML 标记（UTF-8，`%%SECTION%%` 分隔）
├─ installer_template.txt     ← 一键安装器的模板（纯 ASCII）
├─ installer_msgs.txt         ← 安装器弹窗里的中文文案（UTF-8，一行一条）
├─ uninstaller_template.txt   ← 一键卸载器的模板（纯 ASCII）
├─ uninstaller_msgs.txt       ← 卸载器弹窗里的中文文案（UTF-8，一行一条）
├─ make_installer.ps1         ← 把 GMS + 文案 + 工具栏标记内嵌进模板，产出 dist/ 里的安装器与卸载器
└─ PLAN.md                    ← 修复计划 + X4 实测 API 事实（改代码前先看这个）
```

---

## 五、自己改代码 / 重新打包

改 `src/` 里的 `.bas` 就行。**仓库里统一是 UTF-8**（方便在 GitHub 上阅读）；
VBE 导入只认 ANSI，所以 `build_gms.vbs` 会自动把每个文件转成 GBK 临时副本再导入，你不用手动转码。

```powershell
# 1) 重新生成 GMS（需要本机装 X4，脚本会驱动 CorelDRAW 的 VBA 引擎）
cscript //nologo build_gms.vbs

# 2) 重新生成一键安装器 / 卸载器（把新 GMS + 中文文案内嵌进去）
powershell -ExecutionPolicy Bypass -File make_installer.ps1
```

`build_gms.vbs` 的做法：自动在 X4 安装目录里找一个未加密的 GMS（优先 `Emboss.gms`）当种子，
清空组件、导入 `src/` 全部模块，最后保存到
`%APPDATA%\Corel\CorelDRAW Graphics Suite X4\User Draw\GMS\CDRX4Toolkit.gms`。
所有路径都是运行时定位的，不写死盘符。

改工具栏按钮（名称 / 顺序 / 提示）要同时改两处：`src\M_Install.bas` 里的 `CmdList()`，
和 `workspace_markup.txt` 里对应的三段 XML（按钮 `guid` 在两者之间必须一致）。

弹窗文案在 `installer_msgs.txt` / `uninstaller_msgs.txt` 里（一行一条，`\n` 表示换行，
`{N}` / `{LIST}` / `{TB}` 是占位符），改完重跑第 2 步即可。

> 注意：如果机器上已有别的 GMS 也用了 `CDRX4Toolkit` 这个**工程名**，会冲突导致命令找不到。
> 打包 / 分发前先清掉重名工程。

### 怎么自己跑一遍验证

```powershell
# 冒烟：重启 X4，验工程名 / 工具栏中文名与顺序 / 9 个功能是否全 err=0
cscript //nologo build_gms.vbs
cscript //nologo tools\smoke.vbs
```

**冒烟不能证明「启动不崩」**，必须另跑真实启动测试（v1.0.2 的教训）：

```powershell
# 1) 把工具栏标记从工作区剥掉，模拟一台干净机器
powershell -ExecutionPolicy Bypass -File tools\workspace_strip.ps1
# 2) 跑安装器（静默版，不弹窗）
cscript //nologo _test_silent.vbs
# 3) 手动启动 CorelDRW.exe，观察 30 秒进程健康度 + 有没有新的崩溃事件，并截图
powershell -ExecutionPolicy Bypass -File tools\real_launch_test.ps1 -WaitSec 30 -Shot _launch.png
```

第 3 步必须**真的启动 `CORELDRW.EXE`**：`CreateObject` / `InitializeVBA` / `RunMacro` 走的是
CorelDRAW 的另一条启动路径，**证明不了双击启动不崩**。判定标准是 `==== LAUNCH PASS ====`。
日志落在 `_launch.log` / `_smoke.log` / `_build.log`，已归档副本见 `logs/`。

---

## 六、技术要点（改代码前必读）

### 三条红线

1. **不要用启动钩子装工具栏**（v1.0.2 的教训）。
   `GlobalMacroStorage_Start` 触发时 CorelDRAW 的命令栏框架还没初始化完，在里面
   `CommandBars.Add / Delete / Visible` 会让 X4 在 `CrlFrmWk.dll` 里踩空崩溃
   （`0xc0000005` → `0xc000041d`），而且 `On Error Resume Next` **拦不住进程级访问违规**。
   本包现在**完全没有启动钩子**（`src\_startup.txt` = `--none`）。
2. **不要用 CommandBars API 建 / 删工具栏**。
   外部自动化（`CreateObject` + `CommandBars.Add / AddCustomButton`）在「工作区里还没有这条工具栏」时
   同样会崩 X4，症状一模一样，已复现两次。这条路本身不行，不是时机不对。
3. **工作区文件的两个坑**：
   ① 必须**无 BOM**（CorelDRAW 写的就是无 BOM；加 BOM 会让下次启动崩在 `CrlFrmWk.dll`，实测过）；
   ② CorelDRAW **退出时会重写**这个文件，所以安装 / 卸载时必须先关掉 CDR。

### 正确做法：直接写工作区 XML

工具栏在 CorelDRAW 里本来就是纯 XML。安装器写三处，**漏任何一处都会静默失效**：

1. `<items>` 里每个按钮一条
   `<itemData dynamicCommand="CDRX4Toolkit.模块.过程" userCaption="中文名" userToolTip="提示" .../>`
   —— 这一条负责「按钮绑到哪个宏」和中文标题；
2. `<commandBars>` 里一条 `<commandBarData>`：工具栏骨架（`nonLocalizableName="增强工具"`）
   + 按钮 `guidRef` 的**显示顺序**；
3. `<cmdBarLane>` 里一条 `<commandBar guidRef="..." visible="true"/>`
   —— **漏了这条，工具栏会存在但不显示**，这也是最容易误判成「装了个寂寞」的地方。

这份标记与 CorelDRAW 自己在「工具 > 自定义 > 命令 > 宏」里拖拽生成的一模一样，
所以 X4 能自己解析宏路径，**不需要 `AddPluginCommand`**。

安装器还会：改前备份成 `.cdrx4bak`；用 MSXML 校验改完的文本仍是合法 XML 才落盘；
重复运行幂等（先剥旧标记再插，不会出现两条工具栏）；检测到 CorelDRAW 在运行会提醒先关掉。
整个安装约 0.3 秒，且**不启动 CorelDRAW**。

`M_Install.InstallToolbar` 宏仍然保留，但只当兜底：万一工作区写入失败（权限不足 / 工作区结构不符），
用户可以在 CDR 里手动跑一次，或走「工具 > 自定义 > 命令 > 宏」把宏拖到工具栏上 ——
CorelVBA 等成熟开源方案也是这么装的。

### 其它实现约定

- **批量操作**：`Optimization = True` 关屏幕刷新；`BeginCommandGroup / EndCommandGroup` 让整批操作可一次撤销。
- **遍历**：`Document.Pages` → `Page.Shapes.All`；组用 `Shape.Shapes`，容器用 `Shape.PowerClip.Shapes`。
- **尺寸**：所有定位统一先算毫米再换算到文档单位（`MmPerUnit / MmToDoc / PageWHmm`），
  否则会随文档单位（英寸 / 毫米）乱飘。
- **文件编码**：`src/*.bas` 在仓库里是 **UTF-8**（GitHub 上可读），`build_gms.vbs` 导入前自动转 GBK 临时副本
  —— VBE 只认 ANSI，直接导入 UTF-8 会让按钮中文名变乱码。`*_template.txt` / `*_msgs.txt` 是 UTF-8 的构建输入，
  只由 PowerShell 读取；**真正被执行的 `.vbs` 一律保持纯 ASCII**，中文走 `MSG_B64` 运行时解码。

---

## 七、分发

把 `dist` 里的 **`.vbs` 单个文件**发过去就行 —— 插件本体已经以 base64 内嵌在里面，
不用带 GMS，也不用装 VBE、不用导入源码、不用手动找目录。对方双击，重启 CorelDRAW 即可。

对方需要满足：X4 带 VBA 组件；没有装过同名的 `CDRX4Toolkit` 工程。

> 安装器 / 卸载器都是**纯 ASCII** 的。原因：Windows 脚本宿主按系统 ANSI（中文 Windows 即 GBK）读 `.vbs`，
> 文件里若有 UTF-8 编码的中文会被解错，甚至因为多字节序列吃掉引号而直接语法报错。
> 所以中文文案全部以 base64（UTF-8）形式内嵌、运行时解码 —— 文件无论以什么编码保存或传输，行为都一致。

---

## 八、想要现成的开源方案

不想自己维护的话，**蘭雅 CorelVBA** 是目前最成熟的开源 CDR 插件集，Apache 协议，
明确支持 X4～2023（32 / 64 位），覆盖拼版、自动裁切线、智能群组、批量导图、拆字、AI 剪贴板互转、尺寸标注等：

- GitHub：<https://github.com/hongwenjun/corelvba>
- 国内镜像：<https://gitee.com/mirrors/corelvba>

它的分发方式与本包一致：GMS 放进 `Draw\GMS` + 导入工具栏配置（X4 用 `.xslt` 工作区文件）。
值得注意的是，**它也不在启动事件里自动建工具栏**，而是让用户手动把宏拖到工具栏上 ——
这也是本包 v1.0.3 选择「写工作区 XML」而不是「启动时建工具栏」的旁证。

**推荐组合**：日常排版类需求用 CorelVBA；本包补齐它没有的（全部转曲、标准矩形、印章、转CMYK 这类）。

---

## 九、版本历史

| 版本 | 日期 | 要点 |
|---|---|---|
| **v1.2.0** | 2026-09-28 | 「全部转曲」拆出范围选择：弹窗可选 **全部 / 仅文字 / 仅图形**；工具栏仍 9 个按钮。**GMS 变了，已装旧版需重装** |
| **v1.1.0** | 2026-09-28 | 新增一键卸载器 `卸载CDRX4增强工具.vbs`；插件本体与 v1.0.3 逐字节相同，已装 v1.0.3 无需重装 |
| **v1.0.3** | 2026-09-28 | 修掉严重回归：**装上插件后启动 CDR 崩溃**。删掉全部启动钩子 + 工具栏改为写工作区 XML |
| **v1.0.2** | 2026-09-27 | 修掉「点哪个按钮都出错」（4 个模块编译不过）；按钮改中文名；新增自检与冒烟测试 |
| **v1.0.1** | 2026-09-27 | 修掉弹窗中文乱码 `??????`；安装器改纯 ASCII，修掉非中文系统上的「无效字符」报错 |

逐版本的根因分析、实验记录与验证数据见 [CHANGELOG.md](CHANGELOG.md)；
v1.0.3 的完整排查过程（含决定性实验与外部调研）见 [PLAN.md](PLAN.md) §10，卸载器见 §11，转曲范围拆分见 §12。

---

## 十、免责声明与协议

本包为独立重写实现，与「CorelDRAW X4 SP2 精简增强版」及其捆绑的第三方插件
（SecuriDesign、ConverTo、Cachet 等）**无任何关联**，不包含也不分发其代码。
请勿将原包中的商业 GMS 用于再分发。CorelDRAW 为 Corel 公司商标。

MIT License，见 [LICENSE](LICENSE)。可自由使用、修改、再分发，保留版权声明即可。

随附的 `dist/CDRX4Toolkit.gms` 是本仓库源码的编译产物。
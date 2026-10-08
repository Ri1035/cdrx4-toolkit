# 变更记录（CHANGELOG）

版本号规则：`主.次.修订`。**修订号**用于修 bug（不改变对外行为），**次版本号**用于新增功能或改变安装方式。

---

## v1.3.0 — 2026-10-08

**工具栏改「一图标 + 点击展开菜单」**：原来的 9 个平铺按钮收进 1 个「增强工具」按钮，点一下弹出 9 项 flyout 菜单；顺带把「转曲 / 转CMYK / 颜色替换」的对话框与边界逻辑重新打磨。卸载器同步补齐 flyout 清理。

### 变更（工具栏）

- **9 个平铺按钮 → 1 个展开按钮 + flyout 菜单**（`workspace_markup.txt`）。
  主工具栏现在只有 1 个 `itemData`（展开按钮，`bmpCol=0 bmpRow=1031`，**不绑宏**，靠 `flyoutBarRef` 指向菜单栏）；
  菜单栏是 `flyout="true"` 的 `commandBarData`，按原顺序放 9 个功能按钮。
- **纯 XML 实现，不碰 `CommandBars` API**。`flyout` 结构就是 CorelDRAW 自己持久化的格式，
  照写即可；自动化路线（`Delete`/`Add`/`AddCustomButton`）在 X4 上复现过崩溃（`CrlFrmWk.dll`，
  `0xc0000005 → 0xc000041d`，留下 0 线程僵尸），继续回避。
- 按钮名、顺序、9 条宏绑定**完全不变**，`M_Install.CmdList()` 仍是唯一清单来源（展开按钮不占条目）。

### 修复（卸载器 · **必改项**）

- **展开按钮漏删**。展开按钮没有 `dynamicCommand`，旧卸载器只按「宏名含 `CDRX4Toolkit`」和
  「工具栏名叫增强工具」定位，会把它留成**点不动的死按钮**。
  现在安装器与卸载器的 `StripToolkit()` 都增加一条：**`flyoutBarRef` 指向我方 flyout 栏 guid 的
  `itemData` 一并删除**；并按移除的栏 guid 清掉对应的 `<commandBar>` 可见性条目。
- 卸载后工作区里**主栏 / flyout 栏 / 10 个 itemData / lane 条目 / 9 条宏绑定全部清零**（双向验证通过）。

### 变更（三个功能的对话框与操作逻辑）

> 只动对话框语义与边界提示，纯逻辑函数（`*Core`）签名保持兼容。

- **转曲**（`M_Curves.bas`）：范围选择从 `MsgBox(是/否/取消)` 改为 `InputBox` **编号选择**
  （`1=文字+图形` / `2=只转文字` / `3=只转图形`），**「取消」= 放弃**。
  旧做法把「取消」借来表示「仅图形」，用户得先读说明才敢点 —— 这是本轮修掉的主要可用性问题。
  深度选择的「取消」也明确为放弃；新增 `n=0` 时的三条原因提示。
- **转CMYK**（`M_CMYK.bas`）：对话框写清 `是=转CMYK并给纯黑加叠印` / `否=只转CMYK不动叠印` / `取消=不做`，
  并**明确纯黑 = `C0 M0 Y0 K100`**；新增 `n=0` 提示。
- **颜色替换**（`M_Color.bas`）：取色来源扩展为**优先纯色填充、无则退回轮廓色**；
  新增**方向确认** `InputBox`（`1=按第1个为原色` / `2=反过来` / `取消=不做`）、**同色提示**、
  **`n=0` 三条原因提示**；所有提示都带「原色 → 新色」。

### 新增（验证工具）

- `tools/verify_toolbar.ps1` —— 不启动 CorelDRAW，直接拿 `workspace_markup.txt` 校验工作区：
  `-Expect present` 查「10 guid / 2 骨架 / flyout 标记 / 展开按钮接线 / 主栏只引用展开按钮 /
  flyout 按序 9 项 / 可见性 / 9 条宏绑定」，`-Expect absent` 查全清；并交叉校验 GMS 按钮清单。
- `tools/click_at.ps1` —— 自动点击 + 截图，实证「点展开按钮弹出菜单」。
- `tools/capture_cdr.ps1` —— 枚举进程顶层窗口 + `PrintWindow` 抓图，主窗口被遮挡时也能截到工具栏。

### 验证记录（本机，CorelDRAW X4）

| 项 | 结果 |
|---|---|
| 编译 + 自检 `build_gms.vbs` | `BUILD OK`；`M_Test.SelfTest` 全 `err=0`，含 `1b scope text/shape/all` = **1 / 3 / 0**（与期望一致） |
| 真实启动 `real_launch_test.ps1`（35 s） | 新增 Application Error **0**，进程 `Responding=True` → **LAUNCH PASS**（跑两次均 PASS） |
| 装卸往返 · 卸载 | 两个工作区文件均 **VERIFY PASS (absent)**，展开按钮无残留 |
| 装卸往返 · 安装 | **VERIFY PASS (present)**，flyout 接线与 9 条宏绑定全部正确 |
| **持久性** | 真实启动 → 正常退出 CorelDRAW → 再验仍 **present**，证明 flyout 工具栏被正常持久化 |
| 菜单展开 | 截图 `_our1_crop.png`：点「增强工具」展开 9 项菜单，图标与中文标题正常 |

> **GMS 变了，已装旧版的设备需要重装** —— 双击一次 `dist\安装CDRX4增强工具.vbs` 即可。
> 产物：`CDRX4Toolkit.gms` 166418 字节（SHA-256 `4B17BD79…AB740B`），
> 安装器 252117 字节，卸载器 18037 字节，两个 `.vbs` 均纯 ASCII。

---

## v1.2.0 — 2026-09-28

**「全部转曲」拆出转曲范围**：同一颗按钮，点下去先问「转谁」，可以只转文字或只转图形。

### 新增

- **转曲范围选择**（`src\M_Curves.bas`）。「全部转曲」由原来的「一步问深度」改为**两步弹窗**：
  1. **范围**：是 = 文字 + 图形（全部）／否 = 只转文字／取消 = 只转图形
  2. **深度**：是否深入 PowerClip 容器内部（取消 = 直接返回，什么也不做）

  第一步的「取消」被借来表示第三种范围，所以**第二步的「取消」才是真正的放弃操作**。
- **「图形」= 所有非文本矢量对象**：矩形 / 椭圆 / 多边形 / 符号实例 / 自定义形状 / 连接线……
  按 `Shape.Type` 判定，只排除两类 —— 文本（`cdrTextShape`，由范围决定）和
  本来就是曲线的对象（`cdrCurveShape`，转了也是白转）。位图等转不了的对象
  靠 `ConvertToCurves` 报错跳过，不计入个数。
- 组对象与 PowerClip 的递归遍历对三种范围同样生效。

### 未改动

- **工具栏仍是 9 个按钮**，按钮名、顺序、guid 全部不变。只把「全部转曲」的提示文案改成
  「所有页面的文本与图形转换为曲线（可选只转文字或只转图形）」，
  `M_Install.CmdList` 与 `workspace_markup.txt` 两处同步，安装器重新生成。
- `ConvertAllToCurvesCore` 保持旧签名兼容：`scope` 是**可选**参数，省略即按「全部」处理，
  所以 `M_Test` 里原有的 `ConvertAllToCurvesCore(True)` 调用无需改动。

### 验证记录（本机，CorelDRAW X4）

| 项 | 结果 |
|---|---|
| 自检新增用例 `1b` | 临时文档造 1 个文字 + 矩形 / 椭圆 / 多边形各 1 个，三种范围分三步互相印证：仅文字 → `n=1`、仅图形 → `n=3`（上一步转出的曲线被跳过）、全部 → `n=0`（此时已全是曲线），**全部符合预期** |
| 9 功能冒烟 | 全部 `err=0`（`logs\smoke.log`） |
| 真实启动 | 双击启动 CorelDRAW，30 秒内 0 崩溃事件、进程响应正常 → **LAUNCH PASS** |
| 安装器 | 重新生成，内嵌 GMS SHA-256 `6B7872C4…37F876`（168978 字节），与已安装文件逐字节一致 |

> **GMS 变了，已装旧版的设备需要重装** —— 双击一次 `dist\安装CDRX4增强工具.vbs` 即可。

---

## v1.1.0 — 2026-09-28

**新增一键卸载器** `dist\卸载CDRX4增强工具.vbs`。起因：装了 v1.0.2 的设备上 CDR 会卡死，需要先把旧版清干净。

### 新增
- **`卸载CDRX4增强工具.vbs`**（双击即用，单文件，纯 ASCII + base64 中文）
  1. 删掉所有 X4 目录下的 `CDRX4Toolkit.gms`，含被手工改名成 `.bak` 的；
  2. 把「增强工具」工具栏标记从工作区文件里清掉，改前备份成 `.uninstallbak`。
- **不认版本号**：按「`dynamicCommand` 里含 `CDRX4Toolkit`」和「工具栏名叫 `增强工具`」定位，所以 v1.0.2 及更早版本留下的工具栏一样能清掉。
- **安全闸门**：检测到 CorelDRAW 真的在运行（`ThreadCount > 0`）就直接拒绝执行 —— CDR 退出时会用自己的设置覆盖工作区文件，边运行边改会被冲掉。0 线程的崩溃僵尸不算「在运行」，否则最需要它的那台机器反而用不了。
- 工作区改动前先校验 XML 合法（走临时文件，尊重文件头的 encoding 声明），不合法就不落盘；写回时保持无 BOM。

### 未改动
- **插件本体（GMS）与 v1.0.3 逐字节相同**（SHA-256 `403F8B2E…04CA`，163858 字节）。已装 v1.0.3 的设备**不需要重装**，只是多了一个卸载入口。
- `make_installer.ps1` 扩展为同时产出安装器与卸载器；新增 `uninstaller_template.txt` / `uninstaller_msgs.txt`。

### 验证记录（本机，CorelDRAW X4）
| 项 | 结果 |
|---|---|
| 安装态 → 跑卸载器 | GMS 删除成功；两个工作区文件 `dynamicCommand` 计数 9 → **0**、工具栏名计数 1 → **0**；XML 仍合法；无 BOM；`.uninstallbak` 已生成 |
| 卸载 → 重跑安装器 | 恢复为 `items=9`、`bar=1`、无重复属性、XML 合法、无 BOM；`_default` 工作区**回到 390623 字节**（与 v1.0.3 已验证状态逐字节相同） |
| 内嵌文案完整性 | 两个 .vbs 的 base64 块解码后与 `.txt` 源文件**逐字符相同**（487 / 423 字符） |
| 生成物编码 | 两个 .vbs 均为纯 ASCII（非 ASCII 字符 0 个） |

---

## v1.0.3 — 2026-09-28

修掉 v1.0.2 的严重回归：**装上插件后启动 CorelDRAW X4 直接崩溃，用户无法使用 CDR**。

### 修复

- **删掉全部启动钩子**（致命问题）
  v1.0.2 在 GMS 文档模块里挂了 `GlobalMacroStorage_Start` /
  `GlobalMacroStorage_OnApplicationStart`，启动时调用 `M_Install.InstallToolbarSilent`
  删了再建「增强工具」工具栏。`Start` 触发时 CorelDRAW 的命令栏框架尚未初始化完，
  从启动回调重入 UI 框架导致 `CrlFrmWk.dll` 访问违规（`0xc0000005` →
  `0xc000041d`），进程直接死，留下 0 线程、`taskkill /F` 杀不掉的僵尸。
  `On Error Resume Next` 拦不住进程级访问违规。
  现在 `src\_startup.txt` = `--none`，`build_gms.vbs` 默认不写任何启动代码。
- **工具栏改用「安装器一次性写工作区 XML」**
  不再使用 `CommandBars` API：外部自动化在「工作区里还没有这条工具栏」时同样会崩 X4
  （已复现两次）。安装器现在直接写 CorelDRAW 自己持久化工具栏用的 XML，
  **全程不碰运行中的 CorelDRAW**。
- **中文按钮名跨重启保留**
  标题写入 `<itemData>` 的 `userCaption` / `userToolTip`，随工作区持久化。
  旧版只在运行时设 `Caption`，重启后按钮变回宏路径。

### 变更

- 安装器不再启动 CorelDRAW（原先会 `CreateObject` + `RunMacro` + `Quit`）。
  安装耗时从「数秒 + 起一次 CDR」降到约 0.3 秒。
- `make_installer.ps1` 新增 `workspace_markup.txt` 内嵌（工具栏三段 XML 标记）。
- 安装器新增：工作区备份（`.cdrx4bak`）、写入前 MSXML 合法性校验、
  幂等重装（先剥旧标记再插）、检测 CorelDRAW 是否在运行并提醒。
- `build_gms.vbs` 新增 `--hook=<文件>` / `--nocheck` 参数；启动钩子默认关闭。

### 新增（验证工具）

- `tools/real_launch_test.ps1` — 真的启动 `CORELDRW.EXE`，采样进程健康度、
  查新的 Application Error 事件、截图。**只在启动时跑的代码必须用这个验。**
- `tools/workspace_strip.ps1` — 剥掉工作区里的工具栏标记，模拟一台干净机器。
- `tools/workspace_inject.ps1` — 原型：往工作区注入工具栏标记（安装器已内置同样逻辑）。

### 验证记录（本机，CorelDRAW X4）

| 项 | 结果 |
|---|---|
| 干净工作区 → 安装器 → 真实启动 30 s | 18 线程 / 115.0 MB / Responding=True / 新增崩溃事件 **0** → `LAUNCH PASS` |
| 工具栏 | 「增强工具」出现，9 个中文按钮，顺序 = 全部转曲 → 转CMYK → 颜色替换 → 标准矩形 → 对象适合路径 → JPG批量导出 → 插入页码 → 日历创建 → 印章制作 |
| 工作区三处标记 | `itemData`=9、`commandBarData`=1、`cmdBarLane` 可见性条目=1，XML 合法，无 BOM |
| 幂等性 | 连跑两次安装器，工作区文件 SHA-256 **完全一致**，无重复条目 |
| 安装耗时 | 296 ms |

### 已知事项

- 9 个按钮的**点击效果**（宏是否被 CorelDRAW 正确解析并执行）沿用 CorelDRAW
  原生宏按钮机制，与「工具 > 自定义 > 命令 > 宏」手动拖拽生成的标记格式一致；
  建议安装后手动点一遍做视觉验收（圆/五角星/日历对齐、导出命名等）。
- 崩溃过的机器上会残留 0 线程僵尸进程，**需重启系统**才会消失；不影响新启动的 CDR。

---

## v1.0.2 — 2026-09-27

把「点哪个按钮都出错」查清了。根因是旧源码里有 4 个模块编译不过 ——
VBA 的规则是「一个模块编译失败 → 整个工程所有宏全废」。

### 修复

- 工具栏按钮不是中文名（`M_Install` 未写 `Caption`）。
- `M_Install` 编译错误「类型不匹配」：`Set app = CorelDRAW` 在 X4 里编译不过，
  改为 `CorelDRAW.成员` 取值 + 晚期绑定。
- `M_Rect.StdRectangle` 编译错误「不能给只读属性赋值」：改晚期绑定写 `Radius*`。
- `M_Color.ReplaceColor` 运行时错误 13：`ActiveSelection` 是 Shape，
  改用 `ActiveSelectionRange`。
- `M_CMYK` 编译错误：`CorelDRAW.CreateCMYKColor`；叠印写在 `Shape.OverprintFill`。
- `M_Seal` 语法错误（`circle` 撞保留字）+ 五角星建不出来（改十边形转曲拉顶点）。
- `M_JPG` 导不出文件：`ExportEx` 五参数且必须调 `ex.Finish`。
- `M_FitPath`：X4 的 `CreateBlend` 对外部调用报类型不匹配，改为沿路径节点插值。
- 尺寸随文档单位乱飘：新增毫米换算。

### 新增

- `M_Test` 自检模块 + `tools/smoke.vbs` 冒烟测试。
- `build_gms.vbs` 构建红线防护（干净会话构建、模块名断言、失败回滚）。

### 引入的回归

- 为「开箱即用」增加了启动钩子 `GlobalMacroStorage_Start`，
  **导致 X4 每次启动崩溃** —— 已在 v1.0.3 修复并移除。

---

## v1.0.1 — 2026-09-27

首次推送到 GitHub（<https://github.com/Ri1035/cdrx4-toolkit>，MIT）。
当时推的是早期文件，功能与文档都还不完整。

## v1.0.0 — 2026-09-27

项目起点：用 VBA（GMS）重写原「精简增强版」里的 9 个功能，替代拷贝第三方商业插件。
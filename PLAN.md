# CDRX4Toolkit 修复计划与工作日志

> 本文件是**跨会话的持久化记忆**。上下文被压缩后，先读这里再动手，不要重新调研。

## 0. 当前状态（2026-09-27）

**根因**：`src/*.bas` 里的 4 个模块仍是旧版 buggy 代码，之前那轮"已修复"从未落盘。
用户在别的机器上装的 GMS 因此编译失败 —— **VBA 里任何一个模块有编译错误，整个工程所有宏都跑不了**，
所以表现为"点哪个按钮都出错"。5 张截图与源码逐行对得上。

## 1. 问题 ↔ 源码 对照表（已核实）

| # | 报错位置 | 报错 | 旧代码 | 修法 |
|---|---|---|---|---|
| 1 | `M_Rect.StdRectangle` | 编译错误：不能给只读属性赋值 | `sh.Rectangle.RadiusUpperLeft = 0`（早期绑定报只读） | 用 `Dim ro As Object` 晚期绑定；`ro.RadiusUpperLeft = 0` |
| 2 | `M_Color.ReplaceColor` | 运行时错误 13 类型不匹配 | `CorelDRAW.ActiveSelection` 返回的是 **Shape**，没有 `.Count`/`Item` | 改用 `CorelDRAW.ActiveSelectionRange` |
| 3 | `M_CMYK.CMYKShape` | 编译错误：缺少函数或变量 | `CreateCMYKColor` 裸调用 + `sh.Fill.OverprintFill` | 用 `CorelDRAW.CreateCMYKColor`；叠印写 `sh.OverprintFill` |
| 4 | `M_Seal.CreateSeal` | 语法错误 | `Dim circle As Shape`（`circle` 撞保留字） | 改名 `outerCircle` |
| 5 | `M_Seal` 星形 | 创建失败 | `CreatePolygon(cx,cy,5,5)` 参数错 | 见 §2 星形做法 |
| 6 | 工具栏 | 按钮非中文名 | `M_Install` 未设 `Caption` | 建按钮时直接 `btn.Caption` |
| 7 | 全局 | 尺寸不对 | 页面尺寸随文档单位（英寸/毫米）变 | 统一 `MmPerUnit()` 换算 |

## 2. X4 实测 API 事实（来自 `_probe*.log`，**不要再重新探测**）

### 形状创建（`Layer`）
- `CreateRectangle2(x, y, w, h)` — (x,y)=**左下角**，w/h=**全宽全高**
- `CreateEllipse2(cx, cy, rx, ry)` — (cx,cy)=**中心**，rx/ry=**半径**
- `CreatePolygon2(cx, cy, radius, sides [, startAngle])` — 参数语义已确认
- `CreatePolygon(a,b,c,d,sides)` — 第 5 个参数才是 sides，第 4 个含义不明，**别用**
- `CreateArtisticTextWide(x, y, text)` — OK
- `CreateStar` — **不存在**
- `Shape.Duplicate` / `Duplicate(dx,dy)` / `Clone` — OK
- `Shape.ConvertToCurves` — **语句**形式（`Set x = sh.ConvertToCurves` 会 [424]）

### 五角星做法（CreatePolygon 无法设尖锐度）
用 `CreatePolygon2(cx, cy, rOuter, 10)` 造**十边形** → `ConvertToCurves` → 把**隔一个的顶点**沿半径方向拉到 `rInner`。
节点位置可写：`Curve.Nodes(j).SetPosition x, y`（`PositionX=` 也可写）。
角度用 `Atn` 自己实现 `Atan2`，从节点相对中心的实际角度反推，无需知道起始角。

### 颜色
- `Color.Type`：RGB=5，CMYK=2
- `Color.ConvertToCMYK` — **语句**（就地转换，返回值是 Nothing）
- `Color.CMYKAssign c,m,y,k` — OK，会把颜色就地变成 CMYK（Type=2）
- `Color.RGBAssign r,g,b` — OK
- `Color.IsSame(other)` — OK（返回 Boolean）
- `Color.CopyAssign other` — OK
- `CorelDRAW.CreateRGBColor(r,g,b)` / `CorelDRAW.CreateCMYKColor(c,m,y,k)` — OK
- `Fill.ApplyUniformFill col` / `Fill.ApplyNoFill` / `Outline.Color.CopyAssign col` / `Outline.Width` — OK
- **叠印写在 Shape 上**：`sh.OverprintFill = True`、`sh.OverprintOutline = True`
  （`sh.Fill.OverprintFill` → [438] 不存在）

### 矩形圆角
- **写** `Rectangle.RadiusUpperLeft/UpperRight/LowerLeft/LowerRight = 0` → 运行时 **OK**
  （但早期绑定会报"只读属性"，所以必须 `Dim ro As Object` 晚期绑定）
- `Rectangle.Radius` 写 → [438]；`EqualCorners` 写 → [450]；`SetPolygonProperties` → [438] 不存在

### 选择集
- `CorelDRAW.ActiveSelection` → **Shape**（有 `.Shapes(1)` / `.Shapes.All`，**没有** `.Count`/`.Item`）
- `CorelDRAW.ActiveSelectionRange` → **ShapeRange**（有 `.Count`、`.Item(i)`、`.Shapes(i)`）✅ 用这个

### 导出（`M_JPG` 的关键）
`Document.ExportEx` 返回 `ExportFilter`，**必须调 `ex.Finish` 才会真正写文件**，且是 **5 个参数**：
```vba
Set opt = CorelDRAW.CreateStructExportOptions
opt.ImageType = cdrRGBColorImage
opt.ResolutionX = dpi : opt.ResolutionY = dpi
opt.AntiAliasingType = cdrNormalAntiAliasing      ' 不是 AntiAliasing
Set pal = CorelDRAW.CreateStructPaletteOptions
pg.Activate
Set ex = doc.ExportEx(fileName, cdrJPEG, cdrCurrentPage, opt, pal)
ex.Finish
```
依据：CorelDRAW 自带 `FileConverter.gms` 源码（`_ref/10_FileConverter__frmFileConverter.txt` 第 592/599 行）。
- `opt.AntiAliasing` → [438] 不存在
- 旧的 4 参数写法 + 不调 `Finish` = 文件根本没生成

### 文本
- `t.Text.Story.Font` / `.Size` / `.Alignment` — OK
- `t.Text.FitToPath shape` — OK（**不是** `t.FitToPath`）
- 对齐实测：**3 = 居中**（0/1 左，2 右）

### 曲线几何（做"适合路径"用）
- `Curve.Length`、`Curve.Nodes.Count`、`Curve.Nodes(i).PositionX/Y`、`Curve.SubPaths`、`Curve.Segments(i).StartNode/EndNode` — OK
- `Curve.GetPointPositionAt` / `GetPointAt` 之类**全部不存在** → 只能按节点/线段端点线性插值

### 混合（CreateBlend）— **不可用**
`Shape.CreateBlend(...)` 无论传 Shape / ShapeRange / 多参，一律 [13] Type mismatch。
`doc.CreateBlend` / `lay.CreateBlend` / `app.CreateBlend` 全都不存在。
→ 「对象适合路径」**放弃混合方案**，改为把对象副本沿路径节点/线段插值排布。

### 其他
- `CorelDRAW.Optimization = True/False`、`CorelDRAW.Refresh`、`doc.BeginCommandGroup/EndCommandGroup` — OK
- `GMSManager.RunMacro "CDRX4Toolkit", "M_Util.HasDocument"` — ✅ 唯一可用的调用形式
- `doc.Unit`：1=英寸 2=英尺 3=毫米 4=厘米 5=像素 6=英里 7=米 8=千米 9=迪多点 11=码 12=派卡
- 参考插件 `ConverTo/ColorReplacer/RectangleFixer/FitObjects/CalendarWizard/Cachet_公章` 全部**加密**，读不到源码

## 3. 功能优先级（用户指定）

1. 全部转曲 `M_Curves`
2. 转CMYK `M_CMYK`
3. 颜色替换 `M_Color`
4. 其余随便：标准矩形、JPG批量导出、印章、适合路径、插入页码、日历

用户补充：**印章/日历/页码这类生成类功能，如果难度过高就先放着以后修**。

## 3.5 编译红线：不要 `Set app = CorelDRAW`（2026-09-27 深夜定位）

**症状**：每次启动 CorelDRAW 都弹一个模态框
`Microsoft Visual Basic / 编译错误: 类型不匹配`，工具栏永远装不上，
自动化脚本一调 `RunMacro` 就永久卡死（模态框阻塞 COM）。

**定位过程**：截图看到 VBE 里高亮的是 `M_Install.InstallCore` 的
`Set app = CorelDRAW` 那一行的 `CorelDRAW` 这个词。

**结论**：X4 的 VBA 里把 `CorelDRAW` 这个**全局对象整体赋给变量**编译不过。
但 `CorelDRAW.成员` 这种**取成员**的写法完全没问题（`CorelDRAW.ActiveDocument`
/ `CorelDRAW.CreateDocument` / `CorelDRAW.CommandBars` / `CorelDRAW.AddPluginCommand`
全都是好的，X4 自带宏和 `_ref/` 里的旧版 M_Install 都这么写）。

所以晚期绑定的正确姿势是「把取回来的**成员**放进 Object 变量」：
```vba
Dim cb As Object
Set cb = CorelDRAW.CommandBars(TOOLBAR_NAME)   ' OK
```
而不是「把 CorelDRAW 本体放进变量」：
```vba
Dim app As Object
Set app = CorelDRAW                            ' ✗ 编译错误：类型不匹配
```

**为什么之前没发现**：VBA 是**按需编译**的。`M_Test.SelfTest` 的调用链
根本走不到 `M_Install`，所以构建脚本第 9 步的 selftest 全绿，却漏掉了这个
编译错误；直到 `ThisDocument` 的启动钩子去调 `M_Install.InstallToolbarSilent`
才炸。**构建脚本第 9 步现在必须额外跑一次 `M_Install.InstallToolbarSilent`**，
否则同类 bug 还会漏网。

**附带修正**：`InstallCore` 改成「先 Delete 旧工具栏再重建」。
旧版工具栏的按钮顺序是 转曲→标准矩形→颜色替换→CMYK→…，与用户要求的新顺序
（转曲→CMYK→颜色替换→…）不同，只刷 Caption 会让按钮名和实际动作错位。

## 3.6 构建红线：目标 GMS 不能在 CorelDRAW 里开着（2026-09-27 深夜定位）

**症状**：冒烟测试报
`InstallToolbarSilent err=-2147220224 未找到指定的模块`，
工具栏永远装不上；`_smoke.log` 里赫然写着
```
comp 12 = M_Install1
```

**根因**：X4 启动时会自动加载 `User Draw\GMS\*.gms`，所以**正在运行的
X4 内存里已经有一份 CDRX4Toolkit 工程**。旧 `build_gms.vbs` 直接
`CreateObject` 挂到那个实例上，按文件名找到的其实是**内存里那份旧工程**
（新拷进去的种子文件从头到尾没被打开过）。于是：

1. 旧工程里的 `M_Install` 正卡在编译错误上 → `VBComponents.Remove` 删不掉
   （[3] comps after cleanup = 2，本该是 1）；
2. 重新导入的 `M_Install` 撞名，VBE **静默改名成 `M_Install1`**；
3. 于是 `M_Install.InstallToolbarSilent` 这条宏路径不存在了 →
   `-2147220224 未找到指定的模块`。

**修法**：换文件的时候必须**没有实例在跑**。
`build_gms.vbs` 现在会先检查运行中的 X4：有未保存的文档就报错退出（不碰用户文件），
没有就 `Quit` 并等到进程真的消失，再删旧文件 / 拷种子 / 重新起 X4，
然后轮询等它把种子工程加载进来。另外：
- 清理组件改成「重试 3 轮 + 断言只剩 1 个（ThisDocument）」，不干净就直接失败退出；
- 导入后逐个模块名核对，出现 `Xxx1` 这种改名立刻失败退出，**不再只是打印警告**；
- 失败时把上一版 `CDRX4Toolkit.gms` 从 `.bak` 还原回去，绝不留下一个编译不过的插件
  让 X4 下次启动时自动加载（那会弹出模态框把 COM 卡死）。

**为什么之前的验收漏了**：`_smoke.log` 里的 `M_Install1` 是靠「列组件名」才看出来的。
另外从外部自动化取工具栏也不行 —— `app.CommandBars(名字)` 在 X4 里对外部客户端
只返回 `Nothing` + `err=13 Type mismatch`，所以工具栏的验收必须由 **VBA 自己**
写日志（新增 `M_Install.DiagToolbar` → `%TEMP%\cdrx4_toolbar.log`），
冒烟脚本再读那个日志跟 `src\M_Install.bas` 的按钮表逐项对。

**冒烟测试也要重启 X4**：只有在一个「从磁盘重新加载 GMS」的干净会话里跑，
才算真正验证了保存出来的文件；挂在构建留下的那个实例上等于自欺欺人。

## 4. 执行步骤

- [x] 定位根因 + 汇总 X4 API 事实
- [x] 按 §2 重写 11 个模块（含 `M_Test` 自检）
- [x] 修复工具栏编译红线（§3.5）
- [x] 修复构建红线（§3.6）：干净会话构建 + 模块名断言 + 工具栏 VBA 自检
- [x] 重建 GMS + 逐模块编译校验
- [x] 冒烟测试 9 个功能 + 9 个按钮中文名 → `==== SMOKE PASS ====`
- [x] 重出安装器（内嵌 GMS 与本机验证版 SHA-256 一致）
- [x] 本地交付汇总 → **等用户确认后一起推 GitHub（版本管理 + 更新记录）**

## 5. 约定

- `src/*.bas` 用 UTF-8 保存（方便看 GitHub），`build_gms.vbs` 会转 GBK 再导入
- 每个功能模块拆成两层：`Xxx()` = 带对话框的入口；`XxxCore(...)` = 纯逻辑，供 `M_Test` 自检调用
- 用户要求：**先做本地**，推送 GitHub 前必须等用户确认
- **晚期绑定铁律**：凡是「不确定是否存在」的成员，一律 `Dim o As Object` 后再调。
  VBA 是「一个模块编译不过 → 整个工程所有宏全废」，早期绑定的一个只读属性就能炸掉全部功能。
  这条已经在 M_Rect（Rectangle.Radius*）、M_Install（CommandBars / AddPluginCommand）、
  M_Curves（PowerClip）、M_CMYK（OverprintFill）上踩过。

## 6. 本轮改动日志（2026-09-27 晚）

| 模块 | 关键改动 |
|---|---|
| `M_Util` | 新增 `MmPerUnit/MmToDoc/PageWHmm`（单位换算）、`ActivePageSafe`、`CenterAt`（中心定位，CenterX/Y 失败退 Move）、`SetTextStyle`、`MakeCenteredText`、`ShapeCurve`；颜色工具 `MakeCMYK/MakeRGB/SameColor/ToCMYK`；`Atan2` |
| `M_Curves` | 递归组 / PowerClip（晚期绑定）；`ConvertToCurves` 语句形式；`XxxCore(deep)` |
| `M_CMYK` | `Color.ConvertToCMYK` 就地转换；`CorelDRAW.CreateCMYKColor` 造色；叠印写 `Shape.OverprintFill`；组递归 |
| `M_Color` | 用 `ActiveSelectionRange` + `Item(i)`（`ActiveSelection` 是 Shape，没有 Count/Item）；组递归；`ReplaceColorCore(src,dst)` |
| `M_Rect` | `Dim ro As Object: Set ro = sh.Rectangle` 晚期绑定后写四个 `Radius*`；`StdRectangleCore()` |
| `M_JPG` | 改成 `doc.ExportEx(f, cdrJPEG, cdrCurrentPage, se, pal)` **五参数 + `ex.Finish`**；选项用 `AntiAliasingType`；`ExportJPGTo(folder, dpi)` |
| `M_PageNo` | 统一毫米 → 文档单位换算；`InsertPageNumberCore()` |
| `M_Calendar` | 同上；`CreateCalendarCore(y, m)` |
| `M_Seal` | `circle` 改名 `outerCircle`；`PageWHmm`；`CorelDRAW.CreateRGBColor`；五角星 = 十边形→转曲→隔顶点拉向圆心；`CreateSealCore(txt1,txt2)` |
| `M_FitPath` | 弃用 `CreateBlend`（X4 一律 [13]），改沿路径节点折线均匀插值；路径非曲线时复制转曲读节点后删除；`FitToPathCore(copies)` |
| `M_Install` | 全部改晚期绑定；按钮顺序改为 转曲→CMYK→颜色替换→其余；`Caption` + `TooltipText` 写中文 |
| `M_Test` | 新增自检模块，跑 9 个 Core，结果写 `%TEMP%\cdrx4_selftest.log` |

## 7. 冒烟测试怎么跑

1. `cscript //nologo build_gms.vbs` 重建 GMS
2. `cscript //nologo tools\smoke.vbs` —— 起一个干净的 X4，
   `GMSManager.RunMacro "CDRX4Toolkit", "M_Test.SelfTest"`，
   然后读 `%TEMP%\cdrx4_selftest.log` 判成败
3. 日志里每行 `序号 功能 n=.. err=..`，**err 必须全是 0**

冒烟一共三关：① 13 个组件名干净（无 `M_Install1` 这类改名）
② `M_Install.DiagToolbar` 建出的 9 个按钮，中文名与顺序和 `CmdList()` 一致
③ `M_Test.SelfTest` 的 9 个功能全 `err=0`。全绿才打印 `==== SMOKE PASS ====`。

## 8. 本轮验证记录（2026-09-27 晚，v1.0.2）

| 项 | 值 |
|---|---|
| GMS 路径 | `%APPDATA%\Corel\CorelDRAW Graphics Suite X4\User Draw\GMS\CDRX4Toolkit.gms` |
| GMS 大小 / SHA-256 | 155666 字节 / `B18456B3C5358AE5BED4E84BE04616E9017B11AD87696F2DC4484597B5BD7A8E` |
| 冒烟结果 | `==== SMOKE PASS ====`（`_smoke.log`，22:46） |
| 组件 | 13 个，名字干净（无 `Xxx1` 改名） |
| 工具栏 | 9 个按钮，中文名/顺序与 `CmdList()` 逐项一致 |
| 自检 | curves 3 / cmyk 4 / color 1 / rect 1 / pageno 1 / calendar 38 / seal 4 / fitpath 8 / jpg 1，**err 全 0** |
| 安装器 | `dist\安装CDRX4增强工具.vbs`，216064 字节，纯 ASCII |
| 安装器回归 | 静默跑一遍，写出的 GMS 与上表 SHA-256 **完全一致**；内嵌中文消息 base64 解码后与原文件**逐字节一致** |

### 功能顺序（用户指定）

全部转曲 → 转CMYK → 颜色替换 → 标准矩形 → 对象适合路径 → JPG批量导出 → 插入页码 → 日历创建 → 印章制作

### 还没做的

- 视觉验收：自动化只证明「不报错、处理了对象」，圆/五角星/日历文字的**对齐与间距**需要人在 CDR 里看一眼

## 9. 发布记录

| 日期 | 事件 | 说明 |
|---|---|---|
| 2026-09-27 | 首次推送 GitHub | 仓库 `Ri1035/cdrx4-toolkit` 建立（MIT），推的是 v1.0.1 时代的旧文件 |
| 2026-09-28 | **v1.0.2 全量推送** | commit `5934c04`，42 个文件 / 493911 字节：`src/` 12 模块、`tools/` 冒烟+20 个探针、`dist/安装CDRX4增强工具.vbs`、`README.md`、`PLAN.md`、构建脚本 |
| 2026-09-28 | 文档补丁 | commit `e469ceb`，README 目录树补上 `tools/probe*.vbs`（仓库有、文档漏列） |

推送方式：本机 **没有 git CLI**，走 GitHub git-data API（blobs → tree → commit → PATCH ref）。
大文件（安装器 216KB）先读字节再 base64，body 全部手工转义成纯 ASCII，避免 PowerShell 编码踩坑。

### 本地交付位置

- 工作副本：`c:\Users\Administrator\Desktop\Trae\6ab8d877d5fe7a26dc6b47cf\cdrx4-toolkit`
- 交付副本：`C:\Users\Administrator\Desktop\CDR\cdrx4-toolkit`（43 个文件，与工作副本逐文件 SHA-256 一致）
- 发给别人只需一个文件：`dist\安装CDRX4增强工具.vbs`

---

# 10. 【严重回归】装上插件后启动 CorelDRAW 直接崩溃（2026-09-28）

> **本节是本次事故的唯一权威记录。上下文被压缩后先读这里，不要重新调研。**
> 目标版本：修掉它 → v1.0.3。

## 10.1 现象（用户报告）

装上 v1.0.2 的插件后启动 CorelDRAW X4，**几秒后整个 CDR 失去响应**，用户完全无法使用 CDR。

## 10.2 实测：不是「卡死」，是「崩溃」

任务管理器里看到的那个进程是**崩溃后的僵尸**：`0 线程 / 0 句柄 / 0 模块 / ~1 MB 工作集`，
`Stop-Process -Force` 和 `taskkill /F` 都报 `Access is denied` **杀不掉**（需重启系统才消失）。
**「0 线程」这一点本身就说明它不是卡死——卡死的进程线程数不为 0。**

真实死因在 Windows 应用程序事件日志（`Application`，来源 `Application Error`，id=1000）：

| 时间（09-28） | 异常码 | 出错模块 | 偏移 |
|---|---|---|---|
| 00:54:27 | `0xc0000005` | unknown | `0x0f67b39a` |
| 00:54:34 | `0xc000041d` | unknown | `0x0f67b39a` |
| 00:54:47 | `0xc0000005` | `CrlFrmWk.dll` | `0x0004c3f9` |
| 00:54:53 | `0xc000041d` | `CrlFrmWk.dll` | `0x0004c3f9` |
| 00:56:28 | `0xc0000005` | unknown | `0x7cb9ff9b` |
| 00:56:32 | `0xc000041d` | unknown | `0x7cb9ff9b` |

- `0xc0000005` = 访问违规（Access Violation）
- `0xc000041d` = `STATUS_FATAL_USER_CALLBACK_EXCEPTION`，**用户回调里发生未处理异常**
- `CrlFrmWk.dll` = Corel Framework，即 Corel 自己的界面框架

「未知模块 + 高地址偏移（`0x7cb9ff9b`）」是 **VBA 运行期回调**的典型特征。
`GlobalMacroStorage_Start` 正是以**回调**形式进入 CorelDRAW 进程的——我们的宏在启动回调里崩了，
异常穿透成 `0xc000041d`，CDR 直接死。同一秒内 `0xc0000005` → `0xc000041d` 成对出现，
是「先踩空内存、再以致命回调异常收场」的标准组合。

## 10.3 对照实验（决定性证据，已做）

| 实验 | 条件 | 结果 |
|---|---|---|
| A | 把 `CDRX4Toolkit.gms` 移走（隔离成 `.bak`），**手动**启动 `CORELDRW.EXE` | **完全正常**：13 线程 / 810 句柄 / 111.9 MB / 标题 `CorelDRAW X4 ( 专业版 ) - [图形1]` / 响应正常；观察 60 s **无任何新崩溃事件** |

→ **插件就是元凶。** 不是机器环境、不是 VBA 组件缺失、不是别的插件。

补充排除项：`%APPDATA%\Corel\CorelDRAW Graphics Suite X4\User Draw\GMS\` 里**只有**
`CDRX4Toolkit.gms` 一个文件；第三方插件（ConverTo / SecuriDesign / Cachet印章 / FitObjects /
ColorReplacer / RectangleFixer / ToJPG / CalendarWizard …）全在
`C:\Program Files (x86)\CorelDRAW X4\Draw\GMS\`。
所以**不存在「和别的插件抢「增强工具」工具栏名」**这条可能。

另注：事件日志里 09-27 18:12 / 18:44 / 18:45 也有 CORELDRW 崩溃，但出错模块与偏移
（`CorelDrw.dll 0x00011c50`、`CrlFrmWk.dll 0x000a566a`）和本次**完全不同**，
且时间早于插件首次构建（`build_gms.vbs` 建于 09-27 18:32，`dist\CDRX4Toolkit.gms` 建于 18:54）。
那批是**我们自己的探针/构建脚本反复 `app.Quit`、强杀 CDR** 造成的，**与本次事故无关**，不要混为一谈。

## 10.4 为什么冒烟测试没抓到（流程漏洞，必须记住）

`tools/smoke.vbs` 第 2 步调用的是 **`M_Install.DiagToolbar`**，
而 `DiagToolbar` 内部**自己显式调用** `InstallCore False`。
所以 `_smoke.log` 里那 9 个按钮，是**显式调用**建出来的——
**它完全不能证明启动钩子跑过，更不能证明启动钩子安全。**

`GlobalMacroStorage_Start` 这条路径**从头到尾没有任何测试覆盖过**。
而且它在自动化会话（`CreateObject` + `InitializeVBA`）里的行为，和用户手动双击启动 CDR
时的行为**不一样**，所以「自动化里不崩」**推不出**「手动启动不崩」。

> **红线（写进流程）：凡是只在「启动时」跑的代码，必须用「真的启动 CORELDRW.EXE」来验，
> 不能用 `RunMacro` / `CreateObject` 验。**

## 10.5 根因（已确认）

启动钩子由 `build_gms.vbs` 第 **242–258** 行写死进 `ThisDocument`：

```vba
Private Sub GlobalMacroStorage_Start()
    On Error Resume Next
    M_Install.InstallToolbarSilent
End Sub
Private Sub GlobalMacroStorage_OnApplicationStart()
    On Error Resume Next
    M_Install.InstallToolbarSilent
End Sub
```

它每次启动都执行这些**重 UI 操作**（`M_Install.InstallCore`）：

1. `CorelDRAW.AddPluginCommand` × 9
2. `CorelDRAW.CommandBars("增强工具").Delete`
3. `CorelDRAW.CommandBars.Add("增强工具")`
4. `cb.Visible = True`
5. `cb.Controls.AddCustomButton` × 9 + 设 `Caption` / `TooltipText`

**`Start` 事件触发时，Corel 的命令栏 / 界面框架尚未初始化完成。**
此时删、建 CommandBar 等于从启动回调里重入 Corel 的 UI 框架 →
在 `CrlFrmWk.dll` 里踩空 → 访问违规 → `0xc000041d` 致命回调异常。

**关键：`On Error Resume Next` 救不了这种情况。**
访问违规是进程级致命异常，不是 VBA 可捕获错误；钩子里那句 `On Error Resume Next` 形同虚设。

次要隐患：`Start` 与 `OnApplicationStart` **同时挂**，工具栏会被建两遍（删一次建一次 ×2），
即使不崩也是不稳定源。

**为什么自动化里没崩**：`CreateObject` 起的实例，VBA 初始化时机与界面构建顺序和双击启动不同，
`Start` 触发时命令栏系统可能已经就绪；双击启动时触发得更早，于是踩空。

## 10.6 决定性实验（已完成，取代原计划的二分定位）

原计划逐条二分「启动期到底哪一步致命」。实际走了一条更直接的路线，结论也更强：

| 实验 | 条件 | 结果 |
|---|---|---|
| B | 无启动钩子的 GMS + 工作区里**已有**工具栏标记 | 手动启动 25 s 正常 |
| C | 无启动钩子的 GMS + 工作区里**没有**工具栏标记 | 手动启动正常，但没有工具栏 |
| D | `CreateObject` + `CommandBars.Delete / Add / Visible / AddCustomButton`（工作区里没有该工具栏） | **崩**：`CrlFrmWk.dll` `0xc0000005` → `0xc000041d`，留 0 线程僵尸。**复现两次** |
| E | `workspace_strip.ps1` 剥掉全部标记 → 直接往 `DRAWUIConfig.xml` 注入三段标记 → **手动启动 `CORELDRW.EXE`** 30 s | **PASS**：18 线程 / 115 MB / Responding=True / 新增崩溃事件 0；工具栏出现、9 个中文按钮顺序正确（`_launch_inject.png`） |

**结论**：

1. 致命的是**「在 CorelDRAW 里用 CommandBars API 建工具栏」这件事本身**：
   启动回调里做会崩（v1.0.2），外部自动化在「该工具栏尚不存在」时做也会崩（实验 D）。
   这条路彻底放弃，不是「时机不对」，是这条路本身不行。
2. **纯写工作区 XML 的路线被实验 E 证明可用**：不需要任何 CorelDRAW 自动化，
   工具栏就能出现、可见、带中文名，且启动不崩。
3. 实验 E 同时回答了原计划的关键未知项 —— **工具栏确实跨重启保留**，
   因为它本来就在工作区文件里，跟 CorelDRAW 怎么启动无关。

## 10.7 修复方案（已采纳并实施）

**方案 1（已实施）：删掉启动钩子 + 安装器直接写工作区 XML**

1. `ThisDocument` 里两个启动事件**全部删除** → 启动期零 UI 操作。
   `src\_startup.txt` = `--none`，`build_gms.vbs` 默认不写启动代码。
2. 安装器 `安装CDRX4增强工具.vbs` **完全不碰 CorelDRAW**，只做两件事：
   写 GMS、往 `DRAWUIConfig.xml` 写工具栏三段标记（`itemData` / `commandBarData` /
   `cmdBarLane` 可见性条目）。实验 D 证明「自动化建工具栏」必须放弃，
   实验 E 证明「写 XML」已经够用。
3. 工具栏跨重启保留 —— 已由实验 E 确认。
4. 兜底：`M_Install.InstallToolbar` 宏 + `Tools > Customization > Commands > Macros`
   手动拖拽（CorelVBA 等社区方案的标准做法），写进安装器失败提示。

**方案 2（不再需要）**：原计划「保留极轻量钩子」。既然纯写 XML 已经够用，
钩子没有任何存在价值 —— 它只会是下一个崩溃源。

## 10.8 外部调研结论（用户要求「优先找开源/社区方案」）

已派研究代理检索，结论**偏保守**，如实记录：

- **没找到**「`GlobalMacroStorage_Start` 里建 CommandBars 导致 X4 崩溃」的权威社区记录。
  只能确认「启动阶段操作 UI/命令栏属高风险时机」，不能援引为已证缺陷。
- **CorelVBA（hongwenjun/corelvba，蘭雅）的官方安装说明**是：把 `.gms` 放进 GMS 目录 →
  重启 CDR → `Tools > Customization` → `Commands` 下拉选 `Macros` → **手动把宏拖到工具栏**。
  即：**成熟开源方案也不在启动事件里自动建工具栏。**
  <https://corelvba.com/index.php?get=set>
- CorelVBA 的 `UI/Toolbar.bas` 确实存在（含 Win32 API 声明、`VB_PredeclaredId = True`），
  但**没有资料显示**它在启动事件里自动建 CommandBars。
- 官方资料确认：工作区（Workspace）是 **XML 结构**，可导出为 **XSLT**；
  但**没查到 X4 的具体路径/扩展名/样例**。
- 有一条相关社区报告：**过大的 GMS 会显著拖慢 Corel 启动**
  <https://coreldraw-sandbox.ideas.aha.io/ideas/CDGS-I-1172>（本包 155 KB，不算大，仅供参考）
- **没查到** X4 VBA 里有可靠的「延迟执行」机制（无 OnTime/Timer 之类），
  所以「让钩子晚点跑」这条路走不通。

**结论：方案 1 与社区实践一致，是首选。**

## 10.9 本次已执行的动作（可回滚）

| 动作 | 说明 |
|---|---|
| 隔离插件 | `%APPDATA%\Corel\CorelDRAW Graphics Suite X4\User Draw\GMS\CDRX4Toolkit.gms` → 改名 `.bak`。**用户现在可以正常用 CDR。** |
| 僵尸进程 | PID 15148 是崩溃残留，`taskkill /F` 报 Access denied，**需重启系统才消失**；不影响使用，但会占一个进程位 |
| 用户可见状态 | CDR 已恢复正常启动（实验 A 验证过） |

## 10.10 待办清单（全部完成）

- [x] §10.6 定位致命步骤 → 结论：CommandBars API 这条路本身不行（实验 D）
- [x] 验证「工具栏是否跨重启保留」→ 保留（实验 E）
- [x] 改 `build_gms.vbs`：启动钩子从 `src\_startup.txt` 读取，默认 `--none`
- [x] 改 `安装CDRX4增强工具.vbs`：不再启动 CDR，改为写工作区 XML
- [x] 重建 GMS + 冒烟 + **新增「真实双击启动不崩」这一条**（`tools/real_launch_test.ps1`）
- [x] 版本号 → v1.0.3，写更新记录（`CHANGELOG.md` + README §九）；README 里
      「启动时自动出现工具栏」「启动钩子」等说法已改
- [ ] 本地验证通过 → 用户确认 → 推送 GitHub

## 10.11 最终验证记录（v1.0.3，2026-09-28）

**端到端模拟一台干净机器**（顺序即实际执行顺序）：

| 步骤 | 命令 | 结果 |
|---|---|---|
| 1 | `tools\workspace_strip.ps1 -Report` | `items=9 userCaption=49 bytes=390623 bom=False` |
| 2 | `tools\workspace_strip.ps1` | `stripped 9 itemData and 1 commandBarData` → `items=0 bytes=387477` |
| 3 | `cscript _test_silent.vbs` | 296 ms；写 GMS ×1，写工作区 ×2（`_default` + `Adobe(R) Illustrator(R)`） |
| 4 | 三处标记核对 | `itemData`=9、`commandBarData`=1、`cmdBarLane` 可见性=1、`userCaption`=9、重复属性=0、XML 合法、`bom=False`、**字节数 390623（与修复前已验证状态逐字节相同）** |
| 5 | 再跑一次安装器（幂等性） | 工作区 SHA-256 **完全一致**，条目数不变 |
| 6 | `tools\real_launch_test.ps1 -WaitSec 30` | `==== LAUNCH PASS ====`：18 线程 / 811 句柄 / 115.0 MB / Responding=True / **新增崩溃事件 0**；截图 `_launch_v103_final.png` 里「增强工具」工具栏 9 个中文按钮齐全、顺序正确 |

**GMS 内容核对**：`Private Sub GlobalMacroStorage` 命中 **0** 次
（唯一的 1 次 `GlobalMacroStorage_Start` 出现在 `M_Install.bas` 的注释里），
即装到机器上的 GMS **确实不含启动钩子**。

GMS SHA-256：`403F8B2E1E50A6E0BF8F5011A486FF4CB895F1CB86409E588C1F9BB7C00904CA`（163858 字节）

**尚未覆盖的一项**：按钮**点击**后宏是否被执行。依据是：这份标记与 CorelDRAW
自己在「工具 > 自定义 > 命令 > 宏」里拖拽生成的一模一样（`dynamicCategory` 与
内置 `GlobalMacros.*` 宏按钮同属 `2cc24a3e-…`），且 9 条宏路径本身已由冒烟验证存在
（`M_Test.SelfTest` 全 `err=0`）。仍建议人工点一遍做视觉验收。

**遗留环境问题**：崩溃过的机器上会残留 0 线程僵尸 `CorelDRW.exe`，
`Stop-Process -Force` / `taskkill /F` 都报 Access denied，**需重启系统**才消失。
不影响新启动的 CDR，也不影响安装器（安装器用 `Win32_Process.ThreadCount > 0` 判断
「CDR 是否真的在运行」，不会被僵尸误判）。

---

## 11. v1.1.0 新增卸载器（2026-09-28）

### 11.1 为什么需要

装了 v1.0.2 的设备上 CDR 会卡死，要重装就得先把旧版清干净。旧版只留了「手工删 GMS + 右键删工具栏」
两步文档，对着一台 CDR 已经卡死的机器并不好操作 —— 而且工具栏标记还留在工作区里，
删掉 GMS 后会剩下一排点不动的死按钮，看起来像「卸载失败」。

### 11.2 设计约束（沿用 v1.0.3 的红线）

1. **不碰 CorelDRAW**。不 CreateObject CorelDRAW ProgID，不碰 CommandBars ——
   §10.6 实验 D 已证明「用 CommandBars API 建/删工具栏」这条路本身会崩，删也一样危险。
2. **不认版本号**。v1.0.2 的工具栏是 CommandBars API 建的，guid 由 CorelDRAW 随机分配，
   硬编码 guid 根本找不到它。所以定位只用两个稳定特征：
   - `<itemData>` 的 `dynamicCommand` 里含 `CDRX4Toolkit`；
   - `<commandBarData>` 的 `nonLocalizableName` / `userCaption` 是「增强工具」。

   骨架的 guid 从**被删的那个骨架块里现读**，再拿它去删 `<cmdBarLane>` 里对应的可见性条目。
3. **CDR 在跑就拒绝执行**。CDR 退出时会用自己的设置覆盖工作区文件，边运行边改会被冲掉。
   判据用 `ThreadCount > 0`，这样 0 线程的崩溃僵尸不会被误判成「正在运行」——
   否则最需要这台脚本的机器反而跑不起来。
4. **改前校验 + 无 BOM 写回**。XML 合法性走临时文件校验（尊重文件头的 encoding 声明），
   不合法就不落盘；备份到 `.uninstallbak`。

### 11.3 实测（本机，X4）

| 步骤 | 结果 |
|---|---|
| 安装态基线 | `gms=True`，`items=9`，`_default` 390679 字节（CDR 退出时自己重写过一次，原 390623） |
| 跑 `_uninstall_silent.vbs` | GMS 删除成功；两个工作区 `items` 9→**0**、`bar` 1→**0**；XML 仍合法；无 BOM；`.uninstallbak` 已生成 |
| 重跑安装器 | `items=9`、`bar=1`、重复属性 0、XML 合法、无 BOM；`_default` **回到 390623 字节**（与 §10.11 已验证状态逐字节相同）→ 往返无残留 |
| 内嵌文案完整性 | 两个 .vbs 的 base64 解码后与 `.txt` 源**逐字符相同**（487 / 423 字符） |
| 生成物编码 | 两个 .vbs 非 ASCII 字符 **0** 个 |

**GMS 未改动**：SHA-256 `403F8B2E1E50A6E0BF8F5011A486FF4CB895F1CB86409E588C1F9BB7C00904CA`（163858 字节），与 v1.0.3 相同 —— 已装 v1.0.3 的设备不用重装。

### 11.4 待办

- [ ] 用户确认 → 推送 GitHub
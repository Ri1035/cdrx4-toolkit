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
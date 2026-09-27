# CDR X4 增强工具包（CDRX4Toolkit）

**当前版本 v1.0.2**（2026-09-27）· 9 个功能已在干净的 X4 会话里跑过自动化冒烟，全绿（见「更新记录」）

一个可自由分发到**普通 CorelDRAW X4** 的插件包，用 VBA 宏（GMS）实现原「精简增强版」里那批插件的主要功能。

- 全部转曲 / 转CMYK / 颜色替换 / 标准矩形 / 对象适合路径 / JPG批量导出 / 插入页码 / 日历创建 / 印章制作
- 自带工具栏，装好后 CDR 里多出一条「增强工具」工具栏，**启动时自动出现，无需手动配置**
- 全部源码开放，可自行修改、打包、分发
- 仓库：<https://github.com/Ri1035/cdrx4-toolkit>

---

## 一、安装（开箱即用）

**双击 `安装CDRX4增强工具.vbs`** 就行（本地在 `dist\` 里；从 GitHub 用就直接下载这一个文件）。它会自己找到 CorelDRAW X4 的 GMS 目录并把插件写进去，然后弹窗告诉你结果。

装完：

1. 如果 CorelDRAW 正开着，**先关掉它**
2. 重新打开 CorelDRAW
3. 顶部出现「**增强工具**」工具栏，9 个按钮直接可用

安装器做的事就一件：把内嵌的 `CDRX4Toolkit.gms` 写进
`%APPDATA%\Corel\CorelDRAW Graphics Suite X4\User Draw\GMS\`。
它只认 **X4 系列**目录，不会污染 2022 等其它版本。

### 卸载

删掉那个目录下的 `CDRX4Toolkit.gms`，重启 CorelDRAW 即可。
（工具栏会随文件消失而不再重建；若残留，右键工具栏 → 删除。）

### 前提：必须有 VBA 组件

GMS 依赖 VBA。完整版 X4 自带；如果按 `Alt+F11` 打不开 VBA 编辑器，说明 VBA 被裁掉了，需要先补装 VBA 组件，否则本包无法运行。

---

## 二、为什么不用直接拷原插件

原「增强插件」里的 GMS（`ConverTo.gms`、`SecuriDesign.gms`、`Cachet_公章.gms` 等）是**第三方商业插件**，随「精简增强版」捆绑分发。把它们拷到别的机器属于盗版，且部分功能与作者机器环境绑定。

本工具包是**重写实现**，可合法分发。原包里的 `防伪花纹 / 防伪轮廓`（SecuriDesign）是专门的防伪底纹引擎，未包含，也不建议复刻。

---

## 三、目录结构

```
cdrx4-toolkit/
├─ README.md                  ← 本文件
├─ LICENSE                    ← MIT 开源协议
├─ .gitignore
├─ dist/                      ← 直接发这个文件夹给对方
│  ├─ 安装CDRX4增强工具.vbs    ← 双击安装（内嵌插件本体，单文件可用）
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
│  └─ M_Test.bas              ← 自检：跑 9 个功能并写日志，供冒烟测试判成败
├─ tools/                     ← 开发/验证脚本（不参与运行，可删）
│  ├─ smoke.vbs               ← 冒烟测试：重启 X4，验工程 / 工具栏 / 9 个功能
│  └─ probe*.vbs              ← X4 对象模型探针（当初用来核对 API，已归档）
├─ build_gms.vbs              ← 用 src/ 里的源码生成 GMS（需本机装 X4）
├─ installer_template.txt     ← 一键安装器的模板（纯 ASCII）
├─ installer_msgs.txt         ← 安装器弹窗里的中文文案（UTF-8，一行一条）
├─ make_installer.ps1         ← 把 GMS + 文案内嵌进模板，产出 dist/ 里的安装器
└─ PLAN.md                    ← 修复计划 + X4 实测 API 事实（改代码前先看这个）
```

---

## 四、功能说明

工具栏上的按钮顺序即下表顺序（前三个是高频功能，排在前面）：

| 按钮 | 用法 | 说明 |
|---|---|---|
| **全部转曲** | 直接点 | 遍历所有页面的文本与图形转曲；弹窗选是否深入 PowerClip 容器 |
| **转CMYK** | 直接点 | 把 RGB 等填充 / 轮廓转 CMYK；可选纯黑叠印 |
| **颜色替换** | 先选中 **2 个**对象 | 第 1 个 = 原色，第 2 个 = 新色；全文档替换填充与轮廓 |
| **标准矩形** | 先选中矩形 | 把圆角 / 扇形角 / 倒棱角重置为直角 |
| **对象适合路径** | 选中 **1 条曲线 + 2 个对象** | 把对象副本沿路径节点均匀排布（见下） |
| **JPG批量导出** | 直接点 | 按页导出 JPG，可设文件夹与 dpi |
| **插入页码** | 直接点 | 每页底部居中插入「第 X 页 / 共 Y 页」 |
| **日历创建** | 直接点 | 输入 `2026-09`，生成该月文字日历 |
| **印章制作** | 直接点 | 外圆 + 环形文字 + 五角星 + 下部文字 |

### 与原插件的差异

- **转CMYK**：原版的对话框还带 `分辨率 / 反锯齿 / 透明背景 / ICC 配置文件` 等选项，那些是**导出参数**；本版只做色彩模型转换 + 叠印。需要导出型转 CMYK 可以再加。
- **对象适合路径**：X4 的 `CreateBlend` 对外部调用一律报类型不匹配（已实测），所以本版不走混合，改为**按路径节点/线段均匀插值排布对象副本**。结果是散开的独立对象，不是混合组。
- **防伪花纹 / 防伪轮廓**：未实现（专门算法引擎）。

### 需要你实测确认的地方

本轮的 9 个功能已在**干净的 X4 会话**里跑过自动化冒烟（见下），全部 `err=0`；但自动化只验证「不报错、处理了对象」，**视觉效果**请你在 CDR 里各点一遍确认：

1. 圆 / 五角星 / 日历文字的**坐标与对齐**
2. **JPG批量导出**的翻页与文件命名
3. **对象适合路径**的间距是否均匀、朝向是否合适

### 怎么自己跑一遍冒烟

```powershell
cscript //nologo build_gms.vbs     # 重建 GMS（会先退出 CorelDRAW）
cscript //nologo tools\smoke.vbs   # 重启 X4，验工程 / 工具栏 / 9 个功能
```

冒烟会依次检查：13 个组件名是否干净（出现 `M_Install1` 这类改名即失败）、
`M_Install.DiagToolbar` 建出来的 9 个按钮的中文名与顺序是否与源码一致、
`M_Test.SelfTest` 跑 9 个功能是否全 `err=0`。全绿才输出 `==== SMOKE PASS ====`。
日志落在 `_smoke.log`。

---

## 五、自己改代码 / 重新打包

改 `src/` 里的 `.bas` 就行，**仓库里统一是 UTF-8**（方便在 GitHub 上阅读）。VBE 导入只认 ANSI，所以 `build_gms.vbs` 会自动把每个文件转成 GBK 临时副本再导入，你不用手动转码。

```powershell
# 1) 重新生成 GMS（需要本机装 X4，脚本会驱动 CorelDRAW 的 VBA 引擎）
cscript //nologo build_gms.vbs

# 2) 重新生成一键安装器（把新 GMS + 中文文案内嵌进去）
powershell -ExecutionPolicy Bypass -File make_installer.ps1
```

`build_gms.vbs` 的做法：自动在 X4 安装目录里找一个未加密的 GMS（优先 `Emboss.gms`）当种子，清空组件、导入 `src/` 全部模块，再把启动事件写进文档模块，最后保存到
`%APPDATA%\Corel\CorelDRAW Graphics Suite X4\User Draw\GMS\CDRX4Toolkit.gms`。所有路径都是运行时定位的，不写死盘符。

弹窗文案在 `installer_msgs.txt` 里（一行一条，`\n` 表示换行，`{N}` / `{LIST}` / `{TB}` 是占位符），改完重跑第 2 步即可。

> 注意：如果机器上已有别的 GMS 也用了 `CDRX4Toolkit` 这个**工程名**，会冲突导致命令找不到。打包/分发前先清掉重名工程。

---

## 六、分发

把 `dist` 里的 **`安装CDRX4增强工具.vbs` 单个文件**发过去就行——插件本体已经以 base64 内嵌在里面，不用带 GMS，也不用装 VBE、不用导入源码、不用手动找目录。对方双击，重启 CorelDRAW 即可。

对方需要满足：X4 带 VBA 组件；没有装过同名的 `CDRX4Toolkit` 工程。

> 这个安装器是**纯 ASCII** 的。原因：Windows 脚本宿主按系统 ANSI（中文 Windows 即 GBK）读 `.vbs`，文件里若有 UTF-8 编码的中文，会被解错，甚至因为多字节序列吃掉引号而直接语法报错。所以中文文案全部以 base64（UTF-8）形式内嵌在 `MSG_B64` 里、运行时解码——这样文件无论以什么编码保存或传输，行为都一致。

---

## 七、如果你想要现成的开源方案

不想自己维护的话，**蘭雅 CorelVBA** 是目前最成熟的开源 CDR 插件集，Apache 协议，明确支持 X4～2023（32/64 位），覆盖拼版、自动裁切线、智能群组、批量导图、拆字、AI 剪贴板互转、尺寸标注等：

- GitHub：<https://github.com/hongwenjun/corelvba>
- 国内镜像：<https://gitee.com/mirrors/corelvba>

它的分发方式与本包一致：GMS 放进 `Draw\GMS` + 导入工具栏配置（X4 用 `.xslt` 工作区文件）。

**推荐组合**：日常排版类需求用 CorelVBA；本包补齐它没有的（全部转曲、标准矩形、印章、转CMYK 这类）。

---

## 八、技术要点（改代码时看）

- **启动自动装工具栏**：在 GMS 的文档模块里写 `Private Sub GlobalMacroStorage_Start()`，调用 `M_Install.InstallToolbarSilent`。X4 里这个事件名是固定的 `Start`，写成别的名字不会触发。本包同时挂了一个 `GlobalMacroStorage_OnApplicationStart` 作兜底。
- **命令注册**：`CorelDRAW.AddPluginCommand "工程名.模块名.过程名", "按钮名", "提示"`。命令路径写错按钮就点不动；同名命令注册一次后不能再改。
- **工具栏**：`CorelDRAW.CommandBars.Add(名称)` → `.Controls.AddCustomButton cdrCmdCategoryMacros, "工程名.模块名.过程名"`。按钮文字用 `.Caption` 直接写中文，不必等重启刷新。
- **每次都重建**：`InstallCore` 先 `CommandBars(名称).Delete` 再 `Add`，然后按 `CmdList()` 的顺序重新建按钮。之所以不做「存在就跳过」的幂等，是因为按钮**顺序**也是要保证的：早期版本建出来的顺序与现在不同，只刷 `Caption` 会让按钮名和实际动作错位。重建一次的开销可以忽略。
- **批量操作**：`Optimization = True` 关屏幕刷新；`BeginCommandGroup / EndCommandGroup` 让整批操作可一次撤销。
- **遍历**：`Document.Pages` → `Page.Shapes.All`；组用 `Shape.Shapes`，容器用 `Shape.PowerClip.Shapes`。
- **文件编码**：`src/*.bas` 在仓库里是 **UTF-8**（GitHub 上可读），`build_gms.vbs` 导入前自动转成 GBK 临时副本——VBE 只认 ANSI，直接导入 UTF-8 会让按钮中文名变乱码。`installer_template.txt` / `installer_msgs.txt` 是 UTF-8 的构建输入，只由 PowerShell 读取；**真正被执行的 `.vbs` 一律保持纯 ASCII**，中文走 `MSG_B64` 运行时解码。`make_installer.ps1` 自身也是纯 ASCII。

---

## 九、更新记录

### v1.0.2（2026-09-27）

这一版把「点哪个按钮都出错」彻底查清了。根因是**旧版源码里有 4 个模块编译不过**——VBA 的规则是「一个模块编译失败 → 整个工程所有宏全废」，所以表现成所有按钮一起坏。

- **修复：工具栏装上后按钮不是中文名**。`M_Install` 建按钮时没写 `Caption`，用的是宏过程名。现在按 `CmdList()` 直接写中文标题与提示。
- **修复：`M_Install` 编译错误「类型不匹配」**。`Set app = CorelDRAW` 这种「把 `CorelDRAW` 全局对象整体赋给变量」的写法在 X4 里编译不过（VBE 会高亮 `CorelDRAW` 这个词）。已全部改为 `CorelDRAW.成员` 取值、把取回的**成员**放进 `Object` 变量做晚期绑定。这条红线已写进 `PLAN.md §3.5`。
- **修复：`M_Rect.StdRectangle` 编译错误「不能给只读属性赋值」**。`Rectangle.Radius*` 早期绑定会被判只读，改用 `Dim ro As Object` 晚期绑定后正常写入。
- **修复：`M_Color.ReplaceColor` 运行时错误 13**。`CorelDRAW.ActiveSelection` 返回的是 **Shape**（没有 `.Count` / `.Item`），已改用 `ActiveSelectionRange`。
- **修复：`M_CMYK` 编译错误「缺少函数或变量」**。`CreateCMYKColor` 要写 `CorelDRAW.CreateCMYKColor`；叠印属性挂在 **Shape** 上（`Shape.OverprintFill`），不是 `Fill` 上。
- **修复：`M_Seal` 语法错误 + 五角星建不出来**。`circle` 撞 VBA 保留字，改名 `outerCircle`；`CreatePolygon` 参数语义不可靠，改为「十边形 → 转曲 → 隔一个顶点沿半径拉到内径」。
- **修复：`M_JPG` 导不出文件**。`ExportEx` 是**五参数**且**必须调 `ex.Finish`** 才会真正写盘；抗锯齿选项名是 `AntiAliasingType`。
- **修复：`M_FitPath` 建混合失败**。X4 的 `CreateBlend` 对外部调用一律报类型不匹配，已改为沿路径节点均匀插值排布副本。
- **修复：尺寸随文档单位（英寸/毫米）乱飘**。新增 `MmPerUnit / MmToDoc / PageWHmm`，所有定位统一先算毫米再换算到文档单位。
- **新增：`M_Test` 自检模块 + `tools/smoke.vbs` 冒烟测试**。冒烟会重启一个**干净的 X4 会话**（只有从磁盘重新加载 GMS 才算真验证），依次核对 13 个组件名是否干净、9 个按钮的中文名与顺序、9 个功能是否全 `err=0`，全绿才输出 `SMOKE PASS`。
- **新增：构建红线防护**。`build_gms.vbs` 现在会先退出运行中的 CorelDRAW 再换文件（否则 VBE 会去改内存里那份旧工程，把新模块静默改名成 `M_Install1`，宏路径全失效）；组件清理后断言只剩 1 个；导入后逐个核对模块名，出现改名立即失败；失败时从 `.bak` 还原上一版，绝不留一个编译不过的 GMS 让 X4 下次启动时弹模态框。
- **功能排序**：按需求调整为 全部转曲 → 转CMYK → 颜色替换 → 标准矩形 → 对象适合路径 → JPG批量导出 → 插入页码 → 日历创建 → 印章制作。

### v1.0.1（2026-09-27）

- **修复：点开功能后弹窗中文全变成 `??????`**。`src/` 里除 `M_Install.bas` 之外的模块，源码中的中文字面量在早期保存时被丢成了 `?`，编译进 GMS 的自然也是问号。现已全部补回并重新编译。
- **修复：部分机器上双击安装器报「无效字符」（800A0408，行 284）**。早期版本的安装器是 **GBK 编码**的 `.vbs`，在非中文 ANSI 代码页的机器上会直接语法报错。现在安装器是**纯 ASCII**、中文走 `MSG_B64` 运行时解码，与系统代码页无关。

> 升级注意：如果装过旧版，**先关掉 CorelDRAW 再重装，装完再启动**，否则 CorelDRAW 里加载的还是旧 GMS。

---

## 十、免责声明

本包为独立重写实现，与「CorelDRAW X4 SP2 精简增强版」及其捆绑的第三方插件（SecuriDesign、ConverTo、Cachet 等）**无任何关联**，不包含也不分发其代码。请勿将原包中的商业 GMS 用于再分发。CorelDRAW 为 Corel 公司商标。

---

## 十一、开源协议

MIT License，见 [LICENSE](LICENSE)。可自由使用、修改、再分发，保留版权声明即可。

本包为独立重写实现，**不包含**任何第三方商业插件的代码；随附的 `dist/CDRX4Toolkit.gms` 是本仓库源码的编译产物。
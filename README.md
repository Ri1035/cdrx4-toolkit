# CDR X4 增强工具包（CDRX4Toolkit）

> 面向 CorelDRAW X4 的 VBA 宏插件：把 9 个常用排版 / 输出功能做成工具栏按钮，双击单个 .vbs 即可安装，可自由分发。

![VBA](https://img.shields.io/badge/-VBA-1470aa?logo=visualbasic&logoColor=white)
![CorelDRAW](https://img.shields.io/badge/-CorelDRAW-X4-000000)
![VBScript](https://img.shields.io/badge/-VBScript-4D6FD2)
![License](https://img.shields.io/badge/license-MIT-green)

## ✨ 项目简介

CDRX4Toolkit 把 CorelDRAW 精简增强版里一批插件的常用功能，用 VBA 宏（GMS）独立重写，打包成双击即用的单个 `.vbs` 安装器。安装后 CorelDRAW 多出一条「增强工具」工具栏，9 个中文按钮打开即用、无需手动配置。全部源码开放，可自行修改、打包与分发。

## 🚀 特性亮点

| 维度 | 说明 |
| --- | --- |
| 全部转曲 | 遍历所有页面对象转曲，可选全部 / 仅文字 / 仅图形，可深入 PowerClip 容器 |
| 转 CMYK | 把 RGB 等填充 / 轮廓转 CMYK，可选纯黑叠印 |
| 颜色替换 | 选中两个对象，全文档替换原色为新色 |
| 标准矩形 | 把圆角 / 扇形角 / 倒棱角重置为直角 |
| 对象适合路径 | 把对象副本沿路径节点均匀排布 |
| JPG 批量导出 | 按页导出 JPG，可设文件夹与 dpi |
| 插入页码 | 每页底部居中插入「第 X 页 / 共 Y 页」 |
| 日历创建 | 输入年月生成文字日历 |
| 印章制作 | 外圆 + 环形文字 + 五角星 + 下部文字 |

## 📖 安装与使用

1. 先关闭 CorelDRAW X4（需已装 VBA 组件，按 `Alt+F11` 可打开 VBA 编辑器）
2. 双击 `dist\安装CDRX4增强工具.vbs`
3. 重新打开 CorelDRAW，「增强工具」工具栏出现，选中对象后点对应按钮即可

卸载：双击 `dist\卸载CDRX4增强工具.vbs`（同样需先关闭 CorelDRAW）。变更记录见 [CHANGELOG.md](CHANGELOG.md)。

适用场景：日常排版 / 输出中反复出现的转曲、转色、批量导出等高频操作，做成工具栏按钮一键完成。

## 🧰 技术栈

- CorelDRAW X4（安装时需勾选 VBA 组件）
- VBA 宏（GMS）· 单文件安装器（VBScript，纯 ASCII、中文文案运行时解码）
- 许可证：MIT

## 🙏 参考与致谢

- 功能对标参考 CorelDRAW 插件市场同类插件：ConverTo、ColorReplacer、RectangleFixer、FitObjects、CalendarWizard、公章插件（上述插件均已加密、无法读取源码，本项目仅作功能对标参考，未引入其代码）

> 各按钮对应的宏逻辑见源码，本文不做展开。

## 修改记录

> 说明：修改记录仅记录文档维护变更，不计入版本更新。

| 日期 | 修改内容 |
| --- | --- |
| 2026-10-05 | README 精简为项目介绍，补充 topics 标签 |
| 2026-10-05 | README 美化丰富：补充在线体验与参考来源，完善使用说明 |

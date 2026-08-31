# BT RF 测试报告生成工具（BTXMLtoExcelTool）

将 **CMWrun** 蓝牙射频（BT RF）传导测试导出的 XML 结果，一键转换为标准化的
Excel 测试报告。基于 WPF（.NET 10）开发，零第三方依赖，可发布为单个免安装 exe。

---

## 功能概述

- **导入 CMWrun XML**：读取 CMWrun 蓝牙测试导出的 XML 结果文件。
- **解析测试数据**：提取测试摘要（设备型号、编号、测试计划、仪器、时间、
  通过/失败统计、最终判定）以及各测试项（TP 分组、限值、实测值、单位、判定）。
- **生成 Excel 报告**：将实测值按行写入内置的标准报告模板（写入设备列 `H`），
  自动把百分比样式从 `0%` 修正为 `0.00%`，避免整数四舍五入丢失精度。
- **内置模板**：报告模板以 Base64 形式内嵌在程序中，无需随程序附带任何模板文件。
- **零外部依赖**：Excel 读写仅使用 .NET 自带的 `System.IO.Compression`（xlsx 本质是
  zip + XML），不引用任何 NuGet 包。

---

## 界面与使用流程

程序为一个轻量的深色主题窗口，操作分两步：

1. **选择 XML**：点击「选择 XML」，选取 CMWrun 导出的蓝牙测试 XML 文件。
2. **生成 Excel 报告**：点击「生成 Excel 报告」，选择保存位置。
   默认文件名为 `<XML 文件名>_Report.xlsx`。

底部状态栏会实时显示当前状态（已选择文件 / 处理中 / 生成成功 / 错误信息）。
生成完成后会弹出提示，告知报告的保存路径。

---

## 运行环境

- **操作系统**：Windows（x86 / x64 均可运行 x86 版本）。
- **运行方式**：
  - 使用自包含单文件发布版（见下）时，**无需**预装 .NET 运行时，
    单个 exe 拷贝到任意 Windows 机器即可直接运行。
  - 若从源码运行，则需要 **.NET 10 SDK**。

---

## 从源码构建与运行

```bash
# 还原并构建
dotnet build BTXMLtoExcelTool.csproj -c Release

# 直接运行
dotnet run --project BTXMLtoExcelTool.csproj
```

## 发布为单文件 exe（推荐）

项目已配置为 **ReadyToRun + 自包含 + 单文件** 发布，并做了如下优化，
使产物为「真正的单文件」——一个 exe 拷到任何 Windows 都能打开：

- 砍掉多余的卫星语言资源（`SatelliteResourceLanguages`）；
- 去掉 ICU 全球化数据（`InvariantGlobalization`），进一步瘦身；
- 把 native 库一并塞进 exe（`IncludeNativeLibrariesForSelfExtract`），
  运行时自动解压到临时目录；
- 单文件内压缩（`EnableCompressionInSingleFile`）；
- 预编译为本机代码以加快启动（`PublishReadyToRun`）。

发布命令：

```bash
dotnet publish BTXMLtoExcelTool.csproj -c Release -r win-x86 --self-contained true -p:PublishSingleFile=true -o <输出目录>
```

也可直接使用随项目提供的发布配置：
`Properties/PublishProfiles/FolderProfile.pubxml`（Visual Studio「发布」即可）。

发布产物为**单个** `BTXMLtoExcelTool.exe`，无附带 DLL、无卫星语言目录、
无 `icu.dll` 等原生库文件。

---

## 项目结构

```
BTXMLtoExcelTool/
├─ App.xaml / App.xaml.cs          # 应用入口
├─ MainWindow.xaml / .xaml.cs      # 主窗口（两步式操作界面）
├─ XmlParser.cs                    # CMWrun XML 解析（摘要 + 测试项分组）
├─ ExcelFiller.cs                  # 核心：内置模板 → 写入实测值 → 输出 xlsx
├─ ExcelExporter.cs               # 基于外部模板文件的填充实现（备用）
├─ Models/Models.cs               # 数据模型（TestSummary / TestGroup / TestItem）
├─ Converters/Converters.cs       # （占位，转换器实际内联在 XAML.cs 中）
├─ AssemblyInfo.cs                 # WPF 主题信息
├─ 蓝牙Breeze-_2_.ico              # 应用图标
├─ Properties/PublishProfiles/     # 发布配置
└─ BTXMLtoExcelTool.csproj         # 项目文件（含单文件发布优化设置）
```

> 说明：仓库中另有 `BtFiller.csproj`、`BtFiller2.csproj`、`BtTestViewer.csproj`
> 等历史/备用项目文件，当前解决方案（`BTXMLtoExcelTool.slnx`）只引用
> `BTXMLtoExcelTool.csproj`。

---

## 输入 XML 格式说明

程序面向 CMWrun 导出的蓝牙测试 XML，主要读取以下结构：

- `Summary`：
  - `DeviceUnderTest/Type`、`DeviceUnderTest/No`
  - `TestPlan`、`User`、`Instrument/ID`
  - `StartTime`、`EndTime`、`TotalTime`、`FinalVerdict`
  - `TestItemsPassed`、`TestItemsFailed`、`TestItemCount`
- `TestItemList`（可出现在树的任意层级）：
  - `ListContext`：用于识别 TP 编号（如 `BV-01-C`）与 TP 名称（如 `Output Power`）
  - `TestItem`：`No`、`Description`、`Condition`、`LowLimit`、`HighLimit`、
    `MeasValue`、`Unit`、`VerdictInfo`、`Verdict`

解析后，实测值会按报告模板的行映射写入设备列（`H` 列，即报告中的 `1#`）。

---

## 技术要点

- **无依赖 xlsx 处理**：xlsx 即 zip 容器，程序直接以 `ZipArchive` 读写内部
  worksheet / styles 的 XML，实现读取、单元格写入、数字格式修正，全程零 NuGet。
- **数值处理**：DEVM 等百分比项以小数存储（如 `0.066`），配合 `0.00%` 样式显示为
  `6.60%`；非百分比项保留 4 位小数。
- **深色 WPF 界面**：自定义按钮模板与状态反馈，操作路径清晰（选择 → 生成）。

---

## 常见问题

- **生成失败 / 报错**：请确认所选 XML 确为 CMWrun 蓝牙测试导出的结果文件，
  且结构完整；状态栏与弹窗会显示具体错误信息。
- **单文件首次启动稍慢**：自包含单文件在首次运行时会将内置的 native 库解压到系统
  临时目录，属正常现象，后续启动会更快。

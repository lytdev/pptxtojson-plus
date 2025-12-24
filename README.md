# 🎨 pptxtojson
一个运行在浏览器中，可以将 .pptx 文件转为可读的 JSON 数据的 JavaScript 库。

> 与其他的pptx文件解析工具的最大区别在于：
> 1. 直接运行在浏览器端；
> 2. 解析结果是**可读**的 JSON 数据，而不仅仅是把 XML 文件内容原样翻译成难以理解的 JSON。

# 🙏 感谢
原作者 [pipipi-pikachu](https://github.com/pipipi-pikachu/pptxtojson) 。
> 因为看原作者时间比较忙，反馈的issues和pull requests都没有及时的修复和合并，所以自己fork了一个，并添加了一些功能。
> 会及时跟进`pipipi-pikachu`的项目。

> 最新跟进commit `62fb9503fcc8e049ad8d4730e75040ffd62c627e`

#  PPTX（Office Open XML）中标签和属性的官方定义与含义
## 1. ECMA-376 标准
[官网地址](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
## 2. Microsoft 官方文档

主要资源：
Open XML SDK Documentation（含类与 XML 元素对应关系）

👉 https://learn.microsoft.com/en-us/office/open-xml/
Specific Element References（按命名空间分类）：
[PresentationML (p:)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation)
[DrawingML (a:)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing)
[Common Elements](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml)


## 3. Open XML SDK Productivity Tool（已归档但可用）
微软曾提供可视化工具，可打开 PPTX 并高亮显示 XML 结构。
虽已停止维护，但仍可下载：
https://github.com/OfficeDev/Open-XML-SDK/releases

## 4. 常见命名空间速查

|前缀	|全称	|用途|
|----|----|----|
|p:	|http://schemas.openxmlformats.org/presentationml/2006/main	|幻灯片、演示文稿结构|
|a:	|http://schemas.openxmlformats.org/drawingml/2006/main	|图形、颜色、几何形状（通用）|
|r:	|http://schemas.openxmlformats.org/officeDocument/2006/relationships	|关系引用（如图片、超链接）|
|cp:	|http://schemas.openxmlformats.org/package/2006/metadata/core-properties	|核心文档属性（作者、标题等）|

# 📄 开源协议
MIT License | Copyright © 2025-PRESENT [lytdev](https://github.com/lytdev/pptxtojson-plus)
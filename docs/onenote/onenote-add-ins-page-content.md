# <a name="work-with-onenote-page-content"></a>处理 OneNote 页面内容 

在 OneNote 外接程序 JavaScript API 中，页面内容由以下对象模型表示。

  ![OneNote 页面对象模型图](../../images/OneNoteOM-page.png)

- Page 对象包含一组 PageContent 对象。
- PageContent 对象包含类型为 Outline、Image 或 Other 的内容。
- Outline 对象包含一组 Paragraph 对象。
- Paragraph 对象包含 RichText、Image、Table 或 Other 这些内容类型。

若要创建空的 OneNote 页面，请使用下列方法之一：

- [Section.addPage](../../reference/onenote/section.md#addpagetitle-string)
- [Page.insertPageAsSibling](../../reference/onenote/page.md#insertpageassiblinglocation-string-title-string)

然后使用以下对象中的方法处理页面内容，如 Page.addOutline 和 Outline.appendHtml。 

- [Page](../../reference/onenote/page.md)
- [Outline](../../reference/onenote/outline.md)
- [Paragraph](../../reference/onenote/paragraph.md)

OneNote 页面的内容和结构由 HTML 进行表示。只有一部分 HTML 可用于创建或更新页面内容，如下所述。

## <a name="supported-html"></a>受支持的 HTML

OneNote 外接程序 JavaScript API 支持使用以下 HTML 创建和更新页面内容：

- `<html>`, `<body>`, `<div>`, `<span>`, `<br/>` 
- `<p>`
- `<img>`
- `<a>`
- `<ul>`, `<ol>`, `<li>` 
- `<table>`, `<tr>`, `<td>`
- `<h1>` ... `<h6>`
- `<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`

## <a name="accessing-page-contents"></a>访问页面内容

只可通过 `Page#load` 访问当前活动页的*页面内容*。 若要更改活动页，请调用 `navigateToPage($page)`。

仍可查询任何页的元数据，例如标题。

## <a name="additional-resources"></a>其他资源

- [OneNote JavaScript API 编程概述](onenote-add-ins-programming-overview.md)
- [OneNote JavaScript API 参考](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Rubric Grader 示例](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office 外接程序平台概述](https://dev.office.com/docs/add-ins/overview/office-add-ins)

---
title: Word JavaScript API 要求集
description: ''
ms.date: 11/14/2018
localization_priority: Priority
ms.openlocfilehash: 0607b9788e4678d608f983c245cef1563e6be362
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389485"
---
# <a name="word-javascript-api-requirement-sets"></a>Word JavaScript API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

Word 外接程序可在多个 Office 版本中运行，包括 Office 2016 for Windows 或更高版本、Office for iPad、Office for Mac 和 Office Online。 下表列出了 Word 要求集、支持该要求集的 Office 主机应用程序，以及这些应用程序的内部版本或版本号。

> [!NOTE]
> 对于标记为 Beta 的要求集，请使用 Office 软件的指定版本（或更高版本），并使用 CDN 的 Beta 库：https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。
> 
> 未列为 Beta 的条目已全面发布，可以继续使用生产 CDN 库：https://appsforoffice.microsoft.com/lib/1/hosted/office.js

|  要求集  |   Office 365 for Windows\*  |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  | Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| WordApi 1.3 | 版本 1612（内部版本 7668.1000）或更高版本| 2017 年 3 月，2.22 或更高版本 | 2017 年 3 月，15.32 或更高版本| 2017 年 3 月 ||
| WordApi 1.2  | 2015 年 12 月更新，版本 1601（内部版本 6568.1000）或更高版本 | 2016 年 1 月，1.18 或更高版本 | 2016 年 1 月，15.19 或更高版本| 2016 年 9 月 | |
| WordApi 1.1  | 版本 1509（内部版本 4266.1001）或更高版本| 2016 年 1 月，1.18 或更高版本 | 2016 年 1 月，15.19 或更高版本| 2016 年 9 月 | |

> [!NOTE]
> 通过 MSI 安装的 Office 2016 的内部版本号为 16.0.4266.1001。 此版本只包含 WordApi 1.1 要求集。

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

- [更新频道发布的 Office 365 客户端版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概述](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="whats-new-in-word-javascript-api-13"></a>Word JavaScript API 1.3 的最近更新 

下面介绍了要求集 1.3 中 Word JavaScript API 的新增内容。 

|对象| 最近更新| 说明|要求集| 
|:-----|-----|:----|:----| 
|[application](/javascript/api/word/word.application)|_方法_ > createDocument(base64File: string) | 使用 base64 编码的.docx 文件创建新文档。 只读。|1.3|
|[body](/javascript/api/word/word.body)|_关系_ > lists|获取 body 中的一组 list 对象。只读。|1.3|
|[body](/javascript/api/word/word.body)|_关系_ > parentBody|获取 body 的父正文。例如，表格单元格 body 的父正文可能是标题。只读。|1.3|
|[body](/javascript/api/word/word.body)|_关系_ > parentSection|获取 body 的父节。只读。|1.3|
|[body](/javascript/api/word/word.body)|_关系_ > styleBuiltIn|获取或设置 body 的嵌入样式名称。请对可以在区域设置之间移植的嵌入样式使用此属性。若要使用自定义样式或本地化样式名称，请参阅“style”属性。|1.3|
|[body](/javascript/api/word/word.body)|_关系_ > tables|获取 body 中的一组 table 对象。只读。|1.3|
|[body](/javascript/api/word/word.body)|_关系_ > type|获取 body 的类型。类型可取值为“MainDoc”、“Section”、“Header”、“Footer”或“TableCell”。只读。|1.3|
|[body](/javascript/api/word/word.body)|_方法_ > getRange(rangeLocation: RangeLocation)|获取整个正文或正文的起点/终点，作为一个范围。|1.3|
|[body](/javascript/api/word/word.body)|_方法_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|插入包含指定行数和列数的表格。insertLocation 的可取值为“Start”或“End”。|1.3|
|[breaktype](/javascript/api/word/word.breaktype)|_关系_ > breaks|指定分隔符的形式：分行、分页或分节类型。 只读。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_关系_ > lists|获取 contentControl 中的一组 list 对象。只读。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_关系_ > parentBody|获取 contentControl 的父正文。只读。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_关系_ > parentTable|获取包含 contentControl 的 table。如果未包含在 table 中，则此关系返回空对象。只读。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_关系_ > parentTableCell|获取包含 contentControl 的 tableCell。如果未包含在 tableCell 中，则此关系返回空对象。只读。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_关系_ > styleBuiltIn|获取或设置 contentControl 的嵌入样式名称。请对可以在区域设置之间移植的嵌入样式使用此属性。若要使用自定义样式或本地化样式名称，请参阅“style”属性。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_关系_ > subtype|获取 contentControl 的子类型。对于 RTF 格式 contentControl，子类型的可取值为“RichTextInline”、“RichTextParagraphs”、“RichTextTableCell”、“RichTextTableRow”和“RichTextTable”。只读。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_关系_ > tables|获取 contentControl 中的一组 table 对象。只读。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_方法_ > getRange(rangeLocation: RangeLocation)|获取整个内容控件或内容控件的起点/终点，作为一个范围。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_方法_ > getTextRanges(endingMarks: string, trimSpacing: bool)|使用标点符号和/或其他结束标记获取内容控件中的文本范围。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_方法_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|将包含指定行数和列数的 table 插入 contentControl 中或在其旁边插入。insertLocation 的可取值为“Start”、“End”、“Before”或“After”。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_方法_ > split(delimiters: string, multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|使用分隔符将内容控件拆分为各个子范围。|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_方法_ > getByTypes(types: ContentControlType)|获取具有指定类型和/或子类型的内容控件。|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_方法_ > getFirst()|获取此集合中的第一个内容控件。|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_属性_ > key|获取 customProperty 的键。 只读。 |1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_属性_ > value|获取或设置 customProperty 的值。|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_关系_ > type|获取自定义属性的值类型。 只读。|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_方法_ > delete()|删除相应的自定义属性。|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_属性_ > items|一组 CustomProperty 对象。只读。|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_方法_ > deleteAll()|删除此集合中的所有自定义属性。|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_方法_ > getCount()|获取自定义属性的计数。|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_方法_ > getItem(key: string)|按键获取自定义属性对象（不区分大小写）。|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_方法_ > set(key: string, value: object)|创建或设置自定义属性。|1.3|
|[document](/javascript/api/word/word.document)|_关系_ > properties|获取当前 document 的属性。只读。|1.3|
|[documentCreated](/javascript/api/word/word.documentcreated)|_方法_ > open()|打开相应文档。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > applicationName|获取 document 的应用程序名称。 只读。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > author|获取或设置 document 的作者。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > category|获取或设置 document 的类别。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > comments|获取或设置 document 的注释。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > company|获取或设置 document 的公司。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > format|获取或设置 document 的格式。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > keywords|获取或设置 document 的关键字。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > lastAuthor|获取或设置 document 的上一作者。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > manager|获取或设置 document 的管理者。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > revisionNumber|获取文档的修订号。 只读。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > security|获取文档的安全性。 只读。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > subject|获取或设置 document 的主题。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > template|获取文档的模板。 只读。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > title|获取或设置 document 的标题。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_关系_ > creationDate|获取文档的创建日期。 只读。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_关系_ > customProperties|获取文档的自定义属性的集合。只读。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_关系_ > lastPrintDate|获取文档的上次打印日期。 只读。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_关系_ > lastSaveTime|获取 document 的上次保存日期。 只读。|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_关系_ > parentTable|获取包含嵌入式图像的 table。如果未包含在 table 中，则此关系返回空对象。只读。|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_关系_ > parentTableCell|获取包含嵌入式图像的 tableCell。如果未包含在 tableCell 中，则此关系返回空对象。只读。|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > getNext()|获取下一个嵌入式图像。|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > getRange(rangeLocation: RangeLocation)|获取图片或图片的起点/终点，作为一个范围。|1.3|
|[inlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|_方法_ > getFirst()|获取此集合中的第一个嵌入式图像。|1.3|
|[list](/javascript/api/word/word.list)|_属性_ > id|获取 list 的 ID。只读。|1.3|
|[list](/javascript/api/word/word.list)|_属性_ > levelExistences|检查 list 中是否包含所有 9 个级别。值为 true 表示级别存在，即各个级别至少存在一个列表项。只读。|1.3|
|[list](/javascript/api/word/word.list)|_关系_ > levelTypes|获取 list 中的所有 9 个级别类型。每种类型的可取值为“Bullet”、“Number”或“Picture”。只读。|1.3|
|[list](/javascript/api/word/word.list)|_关系_ > paragraphs|获取 list 中的段落。只读。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > getLevelParagraphs(level: number)|获取列表中指定级别的段落。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > getLevelString(level: number)|以字符串形式获取指定级别的项目符号、编号或图片。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > insertParagraph(paragraphText: string, insertLocation: InsertLocation)|在指定位置插入段落。insertLocation 的可取值为“Start”、“End”、“Before”或“After”。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > setLevelAlignment(level: number, alignment: Alignment)|设置列表中指定级别的项目符号、编号或图片的对齐方式。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > setLevelBullet(level: number, listBullet: ListBullet, charCode: number, fontName: string)|设置列表中指定级别的项目符号格式。如果项目符号为“Custom”，则需要使用 charCode。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > setLevelIndents(level: number, textIndent: float, textIndent: float)|设置列表中指定级别的两种缩进方式。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > setLevelNumbering(level: number, listNumbering: ListNumbering, formatString: object)|设置列表中指定级别的编号格式。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > setLevelStartingNumber(level: number, startingNumber: number)|设置列表中指定级别的起始编号。默认值为 1。|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_属性_ > items|一组 list 对象。只读。|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_方法_ > getById(id: number)|按标识符获取列表。|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_方法_ > getFirst()|获取此集合中的第一个列表。|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_方法_ > getItem(index: number)|按列表对象在集合中的索引获取列表。|1.3|
|[listItem](/javascript/api/word/word.listitem)|_属性_ > level|获取或设置 list 中项的级别。|1.3|
|[listItem](/javascript/api/word/word.listitem)|_属性_ > listString|以字符串形式获取 listItem 的项目符号、编号或图片。只读。|1.3|
|[listItem](/javascript/api/word/word.listitem)|_属性_ > siblingIndex|获取 listItem 相对于同级元素的序号。只读。|1.3|
|[listItem](/javascript/api/word/word.listitem)|_方法_ > getAncestor(parentOnly: bool)|获取列表项目的父级或最近的上级元素（如果父级不存在的话）。|1.3|
|[listItem](/javascript/api/word/word.listitem)|_方法_ > getDescendants(directChildrenOnly: bool)|获取相应列表项目的所有后代列表项目。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_属性_ > isLastParagraph|指明 paragraph 是其父正文内的最后一个段落。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_属性_ > isListItem|检查 paragraph 是否为 listItem。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_属性_ > tableNestingLevel|获取 paragraph 的表级别。如果 paragraph 未包含在 table 中，则此属性返回 0。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_关系_ > list|获取 paragraph 所属的 List。如果 paragraph 未包含在 list 中，则此关系返回空对象。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_关系_ > listItem|获取 paragraph 的 ListItem。如果 paragraph 未包含在 list 中，则此关系返回空对象。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_关系_ > parentBody|获取 paragraph 的父正文。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_关系_ > parentTable|获取包含 paragraph 的 table。如果未包含在 table 中，则此关系返回空对象。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_关系_ > parentTableCell|获取包含 paragraph 的 tableCell。如果未包含在 tableCell 中，则此关系返回空对象。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_关系_ > styleBuiltIn|获取或设置 paragraph 的嵌入样式名称。请对可以在区域设置之间移植的嵌入样式使用此属性。若要使用自定义样式或本地化样式名称，请参阅“style”属性。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > attachToList(listId: number, level: number)|将 paragraph 加入指定级别的现有 list。如果 paragraph 无法加入 list 或已是 listItem，则无法执行此方法。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > detachFromList()|如果此段落是列表项目的话，从列表中移出此段落。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > getNext()|获取下一个段落。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > getPrevious()|获取上一个段落。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > getRange(rangeLocation: RangeLocation)|获取整个段落或段落的起点/终点，作为一个范围。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > getTextRanges(endingMarks: string, trimSpacing: bool)|使用标点符号和/或其他结束标记获取段落中的文本范围。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|插入包含指定行数和列数的表格。insertLocation 的可取值为“Before”或“After”。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > split(delimiters: string, trimDelimiters: bool, trimSpacing: bool)|使用分隔符将段落拆分为多个子范围。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > startNewList()|生成包含此段落的新列表。如果此段落已是列表项目，则无法执行此方法。|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_方法_ > getFirst()|获取此集合中的第一个段落。|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_方法_ > getLast()|获取此集合中的最后一个段落。|1.3|
|[range](/javascript/api/word/word.range)|_属性_ > hyperlink|获取 range 内的第一个超链接，或在 range 内设置超链接。在 range 内设置新的超链接将删除 range 内的所有超链接。请使用换行符 ('\n') 隔开地址部分和可选的位置部分。|1.3|
|[range](/javascript/api/word/word.range)|_属性_ > isEmpty|检查 range 长度是否为零。只读。|1.3|
|[range](/javascript/api/word/word.range)|_关系_ > lists|获取 range 中的一组 list 对象。只读。|1.3|
|[range](/javascript/api/word/word.range)|_关系_ > parentBody|获取 range 的父正文。只读。|1.3|
|[range](/javascript/api/word/word.range)|_关系_ > parentTable|获取包含 range 的 table。如果未包含在 table 中，则此关系返回空对象。只读。|1.3|
|[range](/javascript/api/word/word.range)|_关系_ > parentTableCell|获取包含 range 的 tableCell。如果未包含在 tableCell 中，则此关系返回空对象。只读。|1.3|
|[range](/javascript/api/word/word.range)|_关系_ > styleBuiltIn|获取或设置 range 的嵌入样式名称。请对可以在区域设置之间移植的嵌入样式使用此属性。若要使用自定义样式或本地化样式名称，请参阅“style”属性。|1.3|
|[range](/javascript/api/word/word.range)|_关系_ > tables|获取 range 中的一组 table 对象。只读。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > compareLocationWith(range: Range)|比较此范围与另一范围的位置。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > expandTo(range: Range)|返回从此范围进行任一方向扩展的新范围，以便覆盖另一范围。此范围不变。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > getHyperlinkRanges()|获取相应范围内的超链接子范围。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > getNextTextRange(endingMarks: string, trimSpacing: bool)|使用标点符号和/或其他结束标记获取下一个文本范围。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > getRange(rangeLocation: RangeLocation)|克隆相应范围，或获取该范围的起点/终点作为一个新范围。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > getTextRanges(endingMarks: string, trimSpacing: bool)|使用标点符号和/或其他结束标记获取相应范围中的文本子范围。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|插入包含指定行数和列数的表格。insertLocation 的可取值为“Before”或“After”。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > intersectWith(range: Range)|返回新范围作为此范围与另一范围的交集。此范围不变。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > split(delimiters: string, multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|使用分隔符将相应范围拆分为各个子范围。|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_属性_ > items|一组 range 对象。只读。|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_方法_ > getFirst()|获取此集合中的第一个范围。|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_方法_ > getItem(index: number)|按范围对象在集合中的索引获取范围。|1.3|
|[requestContext](/javascript/api/word/word.requestcontext)|_方法_ > load(object: object, option: object)|使用参数指定的属性和选项填充在 JavaScript 层中创建的代理对象。 |1.3|
|[requestContext](/javascript/api/word/word.requestcontext)|_方法_ > sync()|将请求队列提交到 Word 并返回一个 promise 对象，此对象可用于将其他操作链接起来。|1.3|
|[section](/javascript/api/word/word.section)|_方法_ > getNext()|获取下一节。|1.3|
|[sectionCollection](/javascript/api/word/word.sectioncollection)|_方法_ > getFirst()|获取此集合中的第一节。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > headerRowCount|获取并设置标题行数。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > height|获取 table 的高度（以磅为单位）。只读。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > isUniform|指明所有表行是否一致。只读。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > nestingLevel|获取 table 的嵌套级别。顶级 table 的级别为 1。只读。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > rowCount|获取 table 的行数。只读。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > shadingColor|获取并设置底纹色。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > style|获取或设置 table 的样式名称。请对自定义样式和本地化样式名称使用此属性。若要使用可以在区域设置之间移植的嵌入样式，请参阅“styleBuiltIn”属性。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > styleBandedColumns|获取并设置 table 是否有镶边列。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > styleBandedRows|获取并设置 table 是否有镶边行。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > styleFirstColumn|获取并设置 table 的第一列是否采用特殊样式。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > styleLastColumn|获取并设置 table 的最后一列是否采用特殊样式。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > styleTotalRow|获取并设置 table 的总计行（最后一行）是否采用特殊样式。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > values|以 2D Javascript 数组形式获取并设置 table 中的文本值。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > width|获取并设置 table 的宽度（以磅为单位）。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > font|获取字体。使用此关系可获取并设置字体名称、大小、颜色和其他属性。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > horizontalAlignment|获取并设置 table 中每个单元格的水平对齐方式。可取值为“left”、“centered”、“right”或“justified”。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > paragraphAfter|获取 table 之后的 paragraph。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > paragraphBefore|获取 table 之前的 paragraph。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > parentBody|获取 table 的父正文。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > parentContentControl|获取包含 table 的 contentControl。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > parentTable|获取包含此 table 的 table。如果未包含在 table 中，则此关系返回空对象。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > parentTableCell|获取包含此 table 的 tableCell。如果未包含在 tableCell 中，则此关系返回空对象。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > rows|获取所有表格行。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > styleBuiltIn|获取或设置 table 的嵌入样式名称。请对可以在区域设置之间移植的嵌入样式使用此属性。若要使用自定义样式或本地化样式名称，请参阅“style”属性。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > tables|获取嵌套一级的子 table。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > verticalAlignment|获取并设置 table 中每个单元格的垂直对齐方式。可取值为“top”、“center”或“bottom”。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > addColumns(insertLocation: InsertLocation, columnCount: number, values: string)|使用第一个或最后一个现有列作为模板，将列添加到 table 的开头或结尾。此方法适用于一致的 table。字符串值（若指定）是在新插入的行中进行设置。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > addRows(insertLocation: InsertLocation, rowCount: number, values: string)|使用第一个或最后一个现有行作为模板，将行添加到 table 的开头或结尾。字符串值（若指定）是在新插入的行中进行设置。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > autoFitContents()|自动调整表列，以适应内容的宽度。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > autoFitWindow()|自动调整表列，以适应窗口的宽度。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > clear()|清除表内容。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > delete()|删除整个表格。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > deleteColumns(columnIndex: number, columnCount: number)|删除特定的列。此方法适用于一致的表格。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > deleteRows(rowIndex: number, rowCount: number)|删除特定的行。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > distributeColumns()|将列设置为等宽。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > distributeRows()|平均分配行高度。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > getBorder(borderLocation: BorderLocation)|获取指定边框的边框样式。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > getCell(rowIndex: number, cellIndex: number)|获取指定行和列处的表单元格。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|获取单元格填充（以磅为单位）。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > getNext()|获取下一个表格。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > getRange(rangeLocation: RangeLocation)|获取包含此表格的范围，或包含此表格的开头或结尾的范围。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > insertContentControl()|在表格中插入内容控件。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > insertParagraph(paragraphText: string, insertLocation: InsertLocation)|在指定位置插入段落。insertLocation 的可取值为“Before”或“After”。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|插入包含指定行数和列数的表格。insertLocation 的可取值为“Before”或“After”。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)|使用指定的 searchOptions 在表格对象范围内执行搜索。搜索结果是一组范围对象。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > select(selectionMode: SelectionMode)|选择表格或其开头或结尾位置，然后将 Word UI 导航到相应位置。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|设置单元格填充（以磅为单位）。|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_属性_ > color|以十六进制值或名称的形式获取或设置表边框颜色。|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_属性_ > width|获取或设置表边框的宽度（以磅为单位）。不适用于宽度固定的表边框类型。|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_关系_ > type|获取或设置表边框的类型。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_属性_ > cellIndex|获取单元格在行中的索引。只读。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_属性_ > columnWidth|获取并设置单元格的列宽度（以磅为单位）。此方法适用于一致的 table。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_属性_ > rowIndex|获取行中单元格在表中的索引。只读。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_属性_ > shadingColor|获取或设置单元格的底纹色。按“#RRGGBB”格式或使用颜色名称指定颜色。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_属性_ > value|获取并设置单元格的文本。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_属性_ > width|获取单元格的宽度（以磅为单位）。只读。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_关系_ > body|获取单元格的 body 对象。只读。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_关系_ > horizontalAlignment|获取并设置单元格的水平对齐方式。可取值为“left”、“centered”、“right”或“justified”。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_关系_ > parentRow|获取单元格的父行。只读。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_关系_ > parentTable|获取单元格的父表。只读。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_关系_ > verticalAlignment|获取并设置单元格的垂直对齐方式。可取值为“top”、“center”或“bottom”。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > deleteColumn()|删除包含该单元格的列。此方法适用于一致的表格。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > deleteRow()|删除包含该单元格的行。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > getBorder(borderLocation: BorderLocation)|获取指定边框的边框样式。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|获取单元格填充（以磅为单位）。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > getNext()|获取下一个单元格。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > insertColumns(insertLocation: InsertLocation, columnCount: number, values: string)|使用单元格的列作为模板，在单元格的左侧或右侧添加列。此方法适用于一致的 table。字符串值（若指定）是在新插入的行中进行设置。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > insertRows(insertLocation: InsertLocation, rowCount: number, values: string)|使用单元格的行作为模板，在单元格的上方或下方插入行。字符串值（若指定）是在新插入的行中进行设置。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|设置单元格填充（以磅为单位）。|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_属性_ > items|一组 tableCell 对象。只读。|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_方法_ > getFirst()|获取此集合中的第一个表格单元格。|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_方法_ > getItem(index: number)|按表格单元格对象在集合中的索引获取表格单元格。|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_属性_ > items|table 对象的集合。只读。|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_方法_ > getFirst()|获取此集合中的第一个表格。|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_方法_ > getItem(index: number)|按表格对象在集合中的索引获取表格。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_属性_ > cellCount|获取行单元格数。只读。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_属性_ > isHeader|检查该行是否为标题行。只读。若要设置标题行数，请对 Table 对象使用 HeaderRowCount。只读。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_属性_ > preferredHeight|获取并设置行的首选高度（以磅为单位）。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_属性_ > rowIndex|获取行在其父表中的索引。只读。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_属性_ > shadingColor|获取并设置底纹色。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_属性_ > values|以 1D Javascript 数组的形式获取并设置行中的文本值。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_关系_ > cells|获取单元格。只读。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_关系_ > font|获取字体。使用此关系可获取并设置字体名称、大小、颜色和其他属性。只读。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_关系_ > horizontalAlignment|获取并设置行中每个单元格的水平对齐方式。可取值为“left”、“centered”、“right”或“justified”。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_关系_ > parentTable|获取父表。只读。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_关系_ > verticalAlignment|获取并设置行中单元格的垂直对齐方式。可取值为“top”、“center”或“bottom”。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > clear()|清除行内容。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > delete()|删除整行。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > getBorder(borderLocation: BorderLocation)|获取行中单元格的边框样式。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|获取单元格填充（以磅为单位）。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > getNext()|获取下一行。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > insertRows(insertLocation: InsertLocation, rowCount: number, values: string)|使用此行作为模板来插入行。如果值已指定，请将值插入新行。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)|使用指定的 searchOptions 在行范围内执行搜索。搜索结果是一组范围对象。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > select(selectionMode: SelectionMode)|选择行，然后将 Word UI 导航到相应位置。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|设置单元格填充（以磅为单位）。|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_属性_ > items|tableRow 对象的集合。只读。|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_方法_ > getFirst()|获取此集合中的第一行。|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_方法_ > getItem(index: number)|按表格行对象在集合中的索引获取表格行。|1.3|


## <a name="whats-new-in-word-javascript-api-12"></a>Word JavaScript API 1.2 的最近更新

下面介绍了要求集 1.2 中 Word JavaScript API 的新增内容。 

|对象| 最近更新| 说明|要求集|
|:-----|-----|:----|:----|
|[contentControl](/javascript/api/word/word.contentcontrol)|_方法_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)|将嵌入式图片插入到内容控件中的指定位置。insertLocation 的可取值为“Replace”、“Start”或“End”。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_关系_ > paragraph|获取包含嵌入式图像的父段落。只读。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > delete()|从文档中删除嵌入式图片。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > insertBreak(breakType: BreakType, insertLocation: InsertLocation)|在主文档的指定位置插入分隔符。insertLocation 的可取值为“Before”或“After”。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > insertFileFromBase64(base64File: string, insertLocation: InsertLocation)|在指定位置插入文档。insertLocation 的可取值为“Before”或“After”。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > insertHtml(html: string, insertLocation: InsertLocation)|在指定位置插入 HTML。insertLocation 的可取值为“Before”或“After”。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation|在指定位置插入嵌入式图片。insertLocation 的可取值为“Replace”、“Before”或“After”。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > insertOoxml(ooxml: string, insertLocation: InsertLocation)|在指定位置插入 OOXML。insertLocation 的可取值为“Before”或“After”。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > insertParagraph(paragraphText: string, insertLocation: InsertLocation)|在指定位置插入段落。insertLocation 的可取值为“Before”或“After”。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > insertText(text: string, insertLocation: InsertLocation)|在指定位置插入文本。insertLocation 的可取值为“Before”或“After”。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > select(selectionMode: SelectionMode)|选择嵌入式图片。这会导致 Word 滚动到选定内容。|1.2|
|[range](/javascript/api/word/word.range)|_关系_ > inlinePictures|获取 range 中的一组 inlinePicture 对象。只读。|1.2|
|[range](/javascript/api/word/word.range)|_方法_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)|在指定位置插入图片。insertLocation 的可取值为“Replace”、“Start”、“End”、“Before”或“After”。|1.2|

## <a name="word-javascript-api-11"></a>Word JavaScript API 1.1

Word JavaScript API 1.1 是首版 API。有关 API 的详细信息，请参阅 [Word JavaScript API](/javascript/api/word) 参考主题。 

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 主机和 API 要求](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 外接程序 XML 清单](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)

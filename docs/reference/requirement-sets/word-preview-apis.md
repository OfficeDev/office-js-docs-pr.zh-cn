---
title: Word JavaScript 预览 Api
description: 有关即将推出的 Word JavaScript Api 的详细信息
ms.date: 07/17/2019
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 3a539f0e69db7c4c567b6fda14f30d6d41a420cf
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/19/2019
ms.locfileid: "35805282"
---
# <a name="word-javascript-preview-apis"></a>Word JavaScript 预览 Api

新 Word JavaScript Api 是在 "预览" 中首次引入的, 并且在完成了充分的测试并获取了用户反馈之后, 它们将成为特定的编号要求集的一部分。

> [!NOTE]
> 预览 API 可能会发生变更，不适合在生产环境中使用。 我们建议你仅在测试和开发环境中试用它们。 不要在生产环境或业务关键型文档中使用预览 API。
>
> 若要使用预览 API，你必须引用 CDN 上的 **beta** 库（https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)并且你可能还需要加入 Office 预览体验成员计划才能获得新的 Office 版本。

## <a name="api-list"></a>API 列表

下表列出了当前预览中的 Api。

| Class | 域 | 说明 |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#ondatachanged)|更改内容控件中的数据时发生。 若要获取新文本, 请在处理程序中加载此内容控件。 若要获取旧文本, 请不要加载它。|
||[onDeleted](/javascript/api/word/word.contentcontrol#ondeleted)|删除内容控件时发生。 请勿在处理程序中加载此内容控件, 否则将无法获取其原始属性。|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onselectionchanged)|在内容控件中的选定内容更改时发生。|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentcontrol)|引发事件的对象。 加载此对象以获取其属性。|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventtype)|事件类型。 有关详细信息, 请参阅 "事件"。|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete--)|删除自定义 XML 部件。|
||[deleteAttribute (xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteattribute-xpath--namespacemappings--name-)|从 xpath 标识的元素中删除具有给定名称的属性。|
||[deleteElement (xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteelement-xpath--namespacemappings-)|删除由 xpath 标识的元素。|
||[getXml()](/javascript/api/word/word.customxmlpart#getxml--)|获取自定义 XML 部件的完整 XML 内容。|
||[insertAttribute (xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertattribute-xpath--namespacemappings--name--value-)|将具有给定名称和值的属性插入到由 xpath 标识的元素中。|
||[insertElement (xpath: string, xml: string, namespaceMappings: any, index？: number)](/javascript/api/word/word.customxmlpart#insertelement-xpath--xml--namespacemappings--index-)|在子位置索引处的 xpath 标识的父元素下插入给定的 XML。|
||[查询 (xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query-xpath--namespacemappings-)|查询自定义 XML 部件的 XML 内容。|
||[id](/javascript/api/word/word.customxmlpart#id)|获取自定义 XML 部件的 ID。 只读。|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceuri)|获取自定义 XML 部件的命名空间 URI。 只读。|
||[setXml (xml: string)](/javascript/api/word/word.customxmlpart#setxml-xml-)|设置自定义 XML 部件的完整 XML 内容。|
||[updateAttribute (xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateattribute-xpath--namespacemappings--name--value-)|使用由 xpath 标识的元素的给定名称更新属性的值。|
||[updateElement (xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateelement-xpath--xml--namespacemappings-)|更新由 xpath 标识的元素的 XML。|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add (xml: string)](/javascript/api/word/word.customxmlpartcollection#add-xml-)|向文档中添加新的自定义 XML 部件。|
||[getByNamespace (namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getbynamespace-namespaceuri-)|获取其命名空间匹配给定命名空间的自定义 XML 部件的新作用域内集合。|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getcount--)|获取集合中项的数目。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getitem-id-)|获取基于其 ID 的自定义 XML 部件。 只读。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getitemornullobject-id-)|获取基于其 ID 的自定义 XML 部件。 如果 CustomXmlPart 不存在, 则返回 null 对象。|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|获取此集合中已加载的子项。|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getcount--)|获取集合中项的数目。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitem-id-)|获取基于其 ID 的自定义 XML 部件。 只读。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitemornullobject-id-)|获取基于其 ID 的自定义 XML 部件。 如果 CustomXmlPart 在集合中不存在, 则返回 null 对象。|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitem--)|如果集合仅包含一个项，则此方法返回该项。 否则, 此方法将产生错误。|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitemornullobject--)|如果集合仅包含一个项，则此方法返回该项。 否则, 此方法将返回 null 对象。|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|获取此集合中已加载的子项。|
|[Document](/javascript/api/word/word.document)|[deleteBookmark (name: string)](/javascript/api/word/word.document#deletebookmark-name-)|从文档中删除书签 (如果存在)。|
||[getBookmarkRange (name: string)](/javascript/api/word/word.document#getbookmarkrange-name-)|获取书签的范围。 如果书签不存在, 则引发。|
||[getBookmarkRangeOrNullObject (name: string)](/javascript/api/word/word.document#getbookmarkrangeornullobject-name-)|获取书签的范围。 如果书签不存在, 则返回 null 对象。|
||[customXmlParts](/javascript/api/word/word.document#customxmlparts)|获取文档中的自定义 XML 部件。 只读。|
||[onContentControlAdded](/javascript/api/word/word.document#oncontentcontroladded)|添加内容控件时发生。 在处理程序中运行 context () 以获取新内容控件的属性。|
||[settings](/javascript/api/word/word.document#settings)|获取文档中的加载项设置。 只读。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark (name: string)](/javascript/api/word/word.documentcreated#deletebookmark-name-)|从文档中删除书签 (如果存在)。|
||[getBookmarkRange (name: string)](/javascript/api/word/word.documentcreated#getbookmarkrange-name-)|获取书签的范围。 如果书签不存在, 则引发。|
||[getBookmarkRangeOrNullObject (name: string)](/javascript/api/word/word.documentcreated#getbookmarkrangeornullobject-name-)|获取书签的范围。 如果书签不存在, 则返回 null 对象。|
||[customXmlParts](/javascript/api/word/word.documentcreated#customxmlparts)|获取文档中的自定义 XML 部件。 只读。|
||[settings](/javascript/api/word/word.documentcreated#settings)|获取文档中的加载项设置。 只读。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageformat)|获取嵌入式图像的格式。 只读。|
|[List](/javascript/api/word/word.list)|[getLevelFont (level: 数字)](/javascript/api/word/word.list#getlevelfont-level-)|获取列表中指定级别的项目符号、编号或图片的字体。|
||[getLevelPicture (level: 数字)](/javascript/api/word/word.list#getlevelpicture-level-)|获取列表中指定级别的图片的 base64 编码的字符串表示形式。|
||[resetLevelFont (level: number, resetFontName？: boolean)](/javascript/api/word/word.list#resetlevelfont-level--resetfontname-)|重置列表中指定级别的项目符号、编号或图片的字体。|
||[setLevelPicture (level: number, base64EncodedImage？: string)](/javascript/api/word/word.list#setlevelpicture-level--base64encodedimage-)|设置列表中指定级别的图片。|
|[Range](/javascript/api/word/word.range)|[getBookmarks (includeHidden？: 布尔值, includeAdjacent？: 布尔值)](/javascript/api/word/word.range#getbookmarks-includehidden--includeadjacent-)|获取或覆盖区域中所有书签的名称。 如果书签的名称以下划线字符开头, 则隐藏该书签。|
||[insertBookmark (name: string)](/javascript/api/word/word.range#insertbookmark-name-)|在区域中插入书签。 如果同一名称的书签存在于某个位置, 则首先将其删除。|
|[设置](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete--)|删除 Setting 对象。|
||[key](/javascript/api/word/word.setting#key)|获取设置的键。 只读。|
||[value](/javascript/api/word/word.setting#value)|获取或设置设置的值。|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add (key: string, value: any)](/javascript/api/word/word.settingcollection#add-key--value-)|创建新设置或设置现有设置。|
||[deleteAll ()](/javascript/api/word/word.settingcollection#deleteall--)|删除此加载项中的所有设置。|
||[getCount()](/javascript/api/word/word.settingcollection#getcount--)|获取设置的计数。|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getitem-key-)|按其键 (区分大小写) 获取设置对象。 如果设置不存在, 则引发。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getitemornullobject-key-)|按其键 (区分大小写) 获取设置对象。 如果设置不存在, 则返回 null 对象。|
||[items](/javascript/api/word/word.settingcollection#items)|获取此集合中已加载的子项。|
|[Table](/javascript/api/word/word.table)|[mergeCells (topRow: 数字, firstCell: number, bottomRow: 数字, lastCell: 数字)](/javascript/api/word/word.table#mergecells-toprow--firstcell--bottomrow--lastcell-)|合并第一个和最后一个单元格所绑定的单元格。|
|[TableCell](/javascript/api/word/word.tablecell)|[split (rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split-rowcount--columncount-)|将单元格拆分为指定的行数和列数。|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertcontentcontrol--)|在行上插入内容控件。|
||[merge ()](/javascript/api/word/word.tablerow#merge--)|将行合并到一个单元格中。|

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Word JavaScript API 要求集](word-api-requirement-sets.md)

---
title: Word JavaScript 预览 API
description: 有关即将推出的 Word JavaScript API 的详细信息。
ms.date: 02/01/2022
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: 4ef8bd9897689b354fa7c19ba0d7be7f8fb92be9
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/02/2022
ms.locfileid: "62320156"
---
# <a name="word-javascript-preview-apis"></a>Word JavaScript 预览 API

新的 Word JavaScript API 首先在"预览版"中引入，之后在经过充分测试并获取用户反馈后，成为特定编号要求集的一部分。

[!INCLUDE [Information about using Word preview APIs](../../includes/word-preview-apis-note.md)]
[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>API 列表

下表列出了当前处于预览中的 Word JavaScript API，但仅在 [Word web 版。](#web-only-api-list) 若要查看所有 Word JavaScript API 的完整列表， (预览 API 和以前发布的 API) ，请参阅 [所有 Word JavaScript API](/javascript/api/word?view=word-js-preview&preserve-view=true)。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#onDataChanged)|更改内容控件内的数据时发生。|
||[onDeleted](/javascript/api/word/word.contentcontrol#onDeleted)|删除内容控件时发生。|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onSelectionChanged)|更改内容控件内的选择时发生。|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentControl)|引发事件的对象。|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventType)|事件类型。|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete__)|删除自定义 XML 部件。|
||[deleteAttribute (xpath： string， namespaceMappings： any， name： string) ](/javascript/api/word/word.customxmlpart#deleteAttribute_xpath__namespaceMappings__name_)|从 xpath 标识的元素中删除具有给定名称的属性。|
||[deleteElement (xpath： string， namespaceMappings： any) ](/javascript/api/word/word.customxmlpart#deleteElement_xpath__namespaceMappings_)|删除由 xpath 标识的元素。|
||[getXml () ](/javascript/api/word/word.customxmlpart#getXml__)|获取自定义 XML 部件的完整 XML 内容。|
||[id](/javascript/api/word/word.customxmlpart#id)|获取自定义 XML 部分的 ID。|
||[insertAttribute (xpath： string， namespaceMappings： any， name： string， value： string) ](/javascript/api/word/word.customxmlpart#insertAttribute_xpath__namespaceMappings__name__value_)|向 xpath 标识的元素插入具有给定名称和值的属性。|
||[insertElement (xpath： string， xml： string， namespaceMappings： any， index？： number) ](/javascript/api/word/word.customxmlpart#insertElement_xpath__xml__namespaceMappings__index_)|在位于子位置索引的 xpath 所标识的父元素下插入给定的 XML。|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceUri)|获取自定义 XML 部分的命名空间 URI。|
||[query (xpath： string， namespaceMappings： any) ](/javascript/api/word/word.customxmlpart#query_xpath__namespaceMappings_)|查询自定义 XML 部分的 XML 内容。|
||[setXml (xml： string) ](/javascript/api/word/word.customxmlpart#setXml_xml_)|设置自定义 XML 部件的完整 XML 内容。|
||[updateAttribute (xpath： string， namespaceMappings： any， name： string， value： string) ](/javascript/api/word/word.customxmlpart#updateAttribute_xpath__namespaceMappings__name__value_)|使用 xpath 标识的元素的给定名称更新属性的值。|
||[updateElement (xpath： string， xml： string， namespaceMappings： any) ](/javascript/api/word/word.customxmlpart#updateElement_xpath__xml__namespaceMappings_)|更新 xpath 标识的元素的 XML。|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[添加 (xml：string) ](/javascript/api/word/word.customxmlpartcollection#add_xml_)|向文档添加新的自定义 XML 部件。|
||[getByNamespace (namespaceUri：string) ](/javascript/api/word/word.customxmlpartcollection#getByNamespace_namespaceUri_)|获取其命名空间匹配给定命名空间的自定义 XML 部件的新作用域内集合。|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getCount__)|获取集合中项的数目。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getItem_id_)|获取基于其 ID 的自定义 XML 部件。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getItemOrNullObject_id_)|获取基于其 ID 的自定义 XML 部件。|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|获取此集合中已加载的子项。|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getCount__)|获取集合中项的数目。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItem_id_)|获取基于其 ID 的自定义 XML 部件。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItemOrNullObject_id_)|获取基于其 ID 的自定义 XML 部件。|
||[getOnlyItem () ](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItem__)|如果集合仅包含一个项，则此方法返回该项。|
||[getOnlyItemOrNullObject () ](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItemOrNullObject__)|如果集合仅包含一个项，则此方法返回该项。|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|获取此集合中已加载的子项。|
|[Document](/javascript/api/word/word.document)|[customXmlParts](/javascript/api/word/word.document#customXmlParts)|获取文档中的自定义 XML 部件。|
||[deleteBookmark (name： string) ](/javascript/api/word/word.document#deleteBookmark_name_)|从文档中删除书签（如果存在）。|
||[getBookmarkRange (name： string) ](/javascript/api/word/word.document#getBookmarkRange_name_)|获取书签的范围。|
||[getBookmarkRangeOrNullObject (name： string) ](/javascript/api/word/word.document#getBookmarkRangeOrNullObject_name_)|获取书签的范围。|
||[ignorePunct](/javascript/api/word/word.document#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.document#ignoreSpace)||
||[matchCase](/javascript/api/word/word.document#matchCase)||
||[matchPrefix](/javascript/api/word/word.document#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.document#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.document#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.document#matchWildcards)||
||[onContentControlAdded](/javascript/api/word/word.document#onContentControlAdded)|添加内容控件时发生。|
||[search (searchText： string， searchOptions？： Word.SearchOptions \| { ignorePunct？： boolean ignoreSpace？： boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： boolean matchWholeWord？： boolean matchWildcards？： boolean }) ](/javascript/api/word/word.document#search_searchText__searchOptions_)|使用指定的搜索选项搜索整个文档的范围。|
||[设置](/javascript/api/word/word.document#settings)|获取文档中加载项的设置。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[customXmlParts](/javascript/api/word/word.documentcreated#customXmlParts)|获取文档中的自定义 XML 部件。|
||[deleteBookmark (name： string) ](/javascript/api/word/word.documentcreated#deleteBookmark_name_)|从文档中删除书签（如果存在）。|
||[getBookmarkRange (name： string) ](/javascript/api/word/word.documentcreated#getBookmarkRange_name_)|获取书签的范围。|
||[getBookmarkRangeOrNullObject (name： string) ](/javascript/api/word/word.documentcreated#getBookmarkRangeOrNullObject_name_)|获取书签的范围。|
||[设置](/javascript/api/word/word.documentcreated#settings)|获取文档中加载项的设置。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageFormat)|获取内嵌图像的格式。|
|[列表](/javascript/api/word/word.list)|[getLevelFont (级别：number) ](/javascript/api/word/word.list#getLevelFont_level_)|获取列表中指定级别的项目符号、编号或图片的字体。|
||[getLevelPicture (级别：number) ](/javascript/api/word/word.list#getLevelPicture_level_)|获取列表中指定级别的图片的 base64 编码字符串表示形式。|
||[resetLevelFont (级别： number， resetFontName？： boolean) ](/javascript/api/word/word.list#resetLevelFont_level__resetFontName_)|重置列表中指定级别的项目符号、编号或图片的字体。|
||[setLevelPicture (level： number， base64EncodedImage？： string) ](/javascript/api/word/word.list#setLevelPicture_level__base64EncodedImage_)|设置列表中指定级别的图片。|
|[区域](/javascript/api/word/word.range)|[getBookmarks (includeHidden？： boolean， includeAdjacent？： boolean) ](/javascript/api/word/word.range#getBookmarks_includeHidden__includeAdjacent_)|获取区域内的所有书签或与区域重叠的名称。|
||[insertBookmark (name： string) ](/javascript/api/word/word.range#insertBookmark_name_)|在范围中插入书签。|
|[设置](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete__)|删除 Setting 对象。|
||[key](/javascript/api/word/word.setting#key)|获取设置的键。|
||[value](/javascript/api/word/word.setting#value)|获取或设置设置的值。|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add (key： string， value： any) ](/javascript/api/word/word.settingcollection#add_key__value_)|创建新设置或设置现有设置。|
||[deleteAll () ](/javascript/api/word/word.settingcollection#deleteAll__)|删除此外接程序中的所有设置。|
||[getCount()](/javascript/api/word/word.settingcollection#getCount__)|获取设置计数。|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getItem_key_)|按其键（区分大小写）获取 setting 对象。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getItemOrNullObject_key_)|按其键（区分大小写）获取 setting 对象。|
||[items](/javascript/api/word/word.settingcollection#items)|获取此集合中已加载的子项。|
|[Table](/javascript/api/word/word.table)|[mergeCells (topRow： number， firstCell： number， bottomRow： number， lastCell： number) ](/javascript/api/word/word.table#mergeCells_topRow__firstCell__bottomRow__lastCell_)|合并第一个单元格和最后一个单元格（包含边界）的单元格。|
|[TableCell](/javascript/api/word/word.tablecell)|[split (rowCount： number， columnCount： number) ](/javascript/api/word/word.tablecell#split_rowCount__columnCount_)|将单元格拆分为指定的行数和列数。|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertContentControl__)|在行上插入内容控件。|
||[merge () ](/javascript/api/word/word.tablerow#merge__)|将行合并为一个单元格。|

## <a name="web-only-api-list"></a>仅 Web API 列表

下表列出了当前仅在 Word web 版 预览版中的 Word JavaScript API。 若要查看所有 Word JavaScript API 的完整列表， (预览 API 和以前发布的 API) ，请参阅 [所有 Word JavaScript API](/javascript/api/word?view=word-js-preview&preserve-view=true)。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[正文](/javascript/api/word/word.body)|[endnotes](/javascript/api/word/word.body#endnotes)|获取正文中的尾注集合。|
||[脚注](/javascript/api/word/word.body#footnotes)|获取正文中的脚注集合。|
||[getComments () ](/javascript/api/word/word.body#getComments__)|获取与正文关联的注释。|
||[getReviewedText (changeTrackingVersion？：Word.ChangeTrackingVersion) ](/javascript/api/word/word.body#getReviewedText_changeTrackingVersion_)|根据 ChangeTrackingVersion 选择获取已审阅文本。|
||[type](/javascript/api/word/word.body#type)|获取 body 的类型。|
|[Comment](/javascript/api/word/word.comment)|[authorEmail](/javascript/api/word/word.comment#authorEmail)|获取批注作者的电子邮件。|
||[authorName](/javascript/api/word/word.comment#authorName)|获取批注作者的姓名。|
||[content](/javascript/api/word/word.comment#content)|获取或设置批注的内容为纯文本。|
||[contentRange](/javascript/api/word/word.comment#contentRange)|获取或设置注释线程状态。|
||[creationDate](/javascript/api/word/word.comment#creationDate)|获取批注的创建日期。|
||[delete()](/javascript/api/word/word.comment#delete__)|删除注释及其回复。|
||[getRange()](/javascript/api/word/word.comment#getRange__)|获取批注位于主文档中的范围。|
||[id](/javascript/api/word/word.comment#id)|ID|
||[replies](/javascript/api/word/word.comment#replies)|获取与注释关联的 reply 对象的集合。|
||[reply (replyText： string) ](/javascript/api/word/word.comment#reply_replyText_)|将新回复添加到注释线程的末尾。|
||[已解决](/javascript/api/word/word.comment#resolved)|获取或设置注释线程的状态。|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[getFirst()](/javascript/api/word/word.commentcollection#getFirst__)|获取集合中的第一个注释。|
||[getFirstOrNullObject () ](/javascript/api/word/word.commentcollection#getFirstOrNullObject__)|获取集合中的第一个注释。|
||[getItem(index: number)](/javascript/api/word/word.commentcollection#getItem_index_)|按注释对象在集合中的索引获取该对象。|
||[items](/javascript/api/word/word.commentcollection#items)|获取此集合中已加载的子项。|
|[CommentContentRange](/javascript/api/word/word.commentcontentrange)|[bold](/javascript/api/word/word.commentcontentrange#bold)|获取或设置一个值，该值指示批注文本是否加粗。|
||[hyperlink](/javascript/api/word/word.commentcontentrange#hyperlink)|获取 range 内的第一个超链接，或在 range 内设置超链接。|
||[insertText (text： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.commentcontentrange#insertText_text__insertLocation_)|将文本插入到指定位置。|
||[isEmpty](/javascript/api/word/word.commentcontentrange#isEmpty)|检查 range 长度是否为零。|
||[italic](/javascript/api/word/word.commentcontentrange#italic)|获取或设置一个值，该值指示批注文本是否为 italicized。|
||[strikeThrough](/javascript/api/word/word.commentcontentrange#strikeThrough)|获取或设置一个值，该值指示批注文本是否有删除线。|
||[text](/javascript/api/word/word.commentcontentrange#text)|获取批注区域的文本。|
||[underline](/javascript/api/word/word.commentcontentrange#underline)|获取或设置一个值，该值指示批注文本的下划线类型。|
|[CommentReply](/javascript/api/word/word.commentreply)|[authorEmail](/javascript/api/word/word.commentreply#authorEmail)|获取批注回复作者的电子邮件。|
||[authorName](/javascript/api/word/word.commentreply#authorName)|获取批注回复作者的姓名。|
||[content](/javascript/api/word/word.commentreply#content)|获取或设置批注回复的内容。|
||[contentRange](/javascript/api/word/word.commentreply#contentRange)|获取或设置 commentReply 的内容范围。|
||[creationDate](/javascript/api/word/word.commentreply#creationDate)|获取批注回复的创建日期。|
||[delete()](/javascript/api/word/word.commentreply#delete__)|删除批注回复。|
||[id](/javascript/api/word/word.commentreply#id)|ID|
||[parentComment](/javascript/api/word/word.commentreply#parentComment)|获取此回复的父批注。|
|[CommentReplyCollection](/javascript/api/word/word.commentreplycollection)|[getFirst()](/javascript/api/word/word.commentreplycollection#getFirst__)|获取集合中的第一个批注回复。|
||[getFirstOrNullObject () ](/javascript/api/word/word.commentreplycollection#getFirstOrNullObject__)|获取集合中的第一个批注回复。|
||[getItem(index: number)](/javascript/api/word/word.commentreplycollection#getItem_index_)|按注释答复对象在集合中的索引获取该对象。|
||[items](/javascript/api/word/word.commentreplycollection#items)|获取此集合中已加载的子项。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[endnotes](/javascript/api/word/word.contentcontrol#endnotes)|获取 contentcontrol 中的尾注集合。|
||[脚注](/javascript/api/word/word.contentcontrol#footnotes)|获取 contentcontrol 中的脚注集合。|
||[getComments () ](/javascript/api/word/word.contentcontrol#getComments__)|获取与正文关联的注释。|
||[getReviewedText (changeTrackingVersion？：Word.ChangeTrackingVersion) ](/javascript/api/word/word.contentcontrol#getReviewedText_changeTrackingVersion_)|根据 ChangeTrackingVersion 选择获取已审阅文本。|
|[Document](/javascript/api/word/word.document)|[changeTrackingMode](/javascript/api/word/word.document#changeTrackingMode)|获取或设置 ChangeTracking 模式。|
||[getEndnoteBody () ](/javascript/api/word/word.document#getEndnoteBody__)|获取单个正文中的文档的尾注。|
||[getFootnoteBody () ](/javascript/api/word/word.document#getFootnoteBody__)|获取单个正文中的文档脚注。|
|[NoteItem](/javascript/api/word/word.noteitem)|[body](/javascript/api/word/word.noteitem#body)|表示便笺项目的 body 对象。|
||[delete()](/javascript/api/word/word.noteitem#delete__)|删除便笺项目。|
||[getNext () ](/javascript/api/word/word.noteitem#getNext__)|获取同一类型的下一个便笺项。|
||[getNextOrNullObject () ](/javascript/api/word/word.noteitem#getNextOrNullObject__)|获取同一类型的下一个便笺项。|
||[reference](/javascript/api/word/word.noteitem#reference)|代表主文档中的脚注或尾注引用。|
||[type](/javascript/api/word/word.noteitem#type)|代表便笺项目类型：脚注或尾注。|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#getFirst__)|获取此集合中的第一个便笺项。|
||[getFirstOrNullObject () ](/javascript/api/word/word.noteitemcollection#getFirstOrNullObject__)|获取此集合中的第一个便笺项。|
||[items](/javascript/api/word/word.noteitemcollection#items)|获取此集合中已加载的子项。|
|[Paragraph](/javascript/api/word/word.paragraph)|[endnotes](/javascript/api/word/word.paragraph#endnotes)|获取段落中的尾注集合。|
||[脚注](/javascript/api/word/word.paragraph#footnotes)|获取段落中的脚注集合。|
||[getComments () ](/javascript/api/word/word.paragraph#getComments__)|获取与段落关联的注释。|
||[getReviewedText (changeTrackingVersion？：Word.ChangeTrackingVersion) ](/javascript/api/word/word.paragraph#getReviewedText_changeTrackingVersion_)|根据 ChangeTrackingVersion 选择获取已审阅文本。|
|[区域](/javascript/api/word/word.range)|[endnotes](/javascript/api/word/word.range#endnotes)|获取范围中的尾注集合。|
||[脚注](/javascript/api/word/word.range#footnotes)|获取范围中的脚注集合。|
||[getComments () ](/javascript/api/word/word.range#getComments__)|获取与区域关联的注释。|
||[getReviewedText (changeTrackingVersion？：Word.ChangeTrackingVersion) ](/javascript/api/word/word.range#getReviewedText_changeTrackingVersion_)|根据 ChangeTrackingVersion 选择获取已审阅文本。|
||[insertComment (commentText： string) ](/javascript/api/word/word.range#insertComment_commentText_)|在范围中插入注释。|
||[insertEndnote (insertText？： string) ](/javascript/api/word/word.range#insertEndnote_insertText_)|插入尾注。|
||[insertFootnote (insertText？： string) ](/javascript/api/word/word.range#insertFootnote_insertText_)|插入脚注。|
|[Table](/javascript/api/word/word.table)|[endnotes](/javascript/api/word/word.table#endnotes)|获取 table 中的尾注集合。|
||[脚注](/javascript/api/word/word.table#footnotes)|获取表格中脚注的集合。|
|[TableRow](/javascript/api/word/word.tablerow)|[endnotes](/javascript/api/word/word.tablerow#endnotes)|获取表格行中的尾注集合。|
||[脚注](/javascript/api/word/word.tablerow#footnotes)|获取表格行中的脚注集合。|

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Word JavaScript API 要求集](word-api-requirement-sets.md)

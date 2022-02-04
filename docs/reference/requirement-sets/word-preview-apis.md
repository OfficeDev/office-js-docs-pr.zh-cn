---
title: Word JavaScript 预览 API
description: 有关即将推出的 Word JavaScript API 的详细信息。
ms.date: 02/01/2022
ms.prod: word
ms.localizationpriority: medium
---

# <a name="word-javascript-preview-apis"></a>Word JavaScript 预览 API

新的 Word JavaScript API 首先在"预览版"中引入，之后在经过充分测试并获取用户反馈后，成为特定编号要求集的一部分。

[!INCLUDE [Information about using Word preview APIs](../../includes/word-preview-apis-note.md)]
[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>API 列表

下表列出了当前处于预览中的 Word JavaScript API，但仅在 [Word web 版。](#web-only-api-list) 若要查看所有 Word JavaScript API 的完整列表， (预览 API 和以前发布的 API) ，请参阅 [所有 Word JavaScript API](/javascript/api/word?view=word-js-preview&preserve-view=true)。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondatachanged-member)|更改内容控件内的数据时发生。|
||[onDeleted](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondeleted-member)|删除内容控件时发生。|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onselectionchanged-member)|更改内容控件内的选择时发生。|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-contentcontrol-member)|引发事件的对象。|
||[eventType](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-eventtype-member)|事件类型。|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-delete-member(1))|删除自定义 XML 部件。|
||[deleteAttribute (xpath： string， namespaceMappings： any， name： string) ](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteattribute-member(1))|从 xpath 标识的元素中删除具有给定名称的属性。|
||[deleteElement (xpath： string， namespaceMappings： any) ](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteelement-member(1))|删除由 xpath 标识的元素。|
||[getXml () ](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-getxml-member(1))|获取自定义 XML 部件的完整 XML 内容。|
||[id](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-id-member)|获取自定义 XML 部分的 ID。|
||[insertAttribute (xpath： string， namespaceMappings： any， name： string， value： string) ](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertattribute-member(1))|向 xpath 标识的元素插入具有给定名称和值的属性。|
||[insertElement (xpath： string， xml： string， namespaceMappings： any， index？： number) ](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertelement-member(1))|在位于子位置索引的 xpath 所标识的父元素下插入给定的 XML。|
||[namespaceUri](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-namespaceuri-member)|获取自定义 XML 部分的命名空间 URI。|
||[query (xpath： string， namespaceMappings： any) ](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-query-member(1))|查询自定义 XML 部分的 XML 内容。|
||[setXml (xml： string) ](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-setxml-member(1))|设置自定义 XML 部件的完整 XML 内容。|
||[updateAttribute (xpath： string， namespaceMappings： any， name： string， value： string) ](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateattribute-member(1))|使用 xpath 标识的元素的给定名称更新属性的值。|
||[updateElement (xpath： string， xml： string， namespaceMappings： any) ](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateelement-member(1))|更新 xpath 标识的元素的 XML。|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[添加 (xml：string) ](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-add-member(1))|向文档添加新的自定义 XML 部件。|
||[getByNamespace (namespaceUri：string) ](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getbynamespace-member(1))|获取其命名空间匹配给定命名空间的自定义 XML 部件的新作用域内集合。|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getcount-member(1))|获取集合中项的数目。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getitem-member(1))|获取基于其 ID 的自定义 XML 部件。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getitemornullobject-member(1))|获取基于其 ID 的自定义 XML 部件。|
||[items](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-items-member)|获取此集合中已加载的子项。|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getcount-member(1))|获取集合中项的数目。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getitem-member(1))|获取基于其 ID 的自定义 XML 部件。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getitemornullobject-member(1))|获取基于其 ID 的自定义 XML 部件。|
||[getOnlyItem () ](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getonlyitem-member(1))|如果集合仅包含一个项，则此方法返回该项。|
||[getOnlyItemOrNullObject () ](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getonlyitemornullobject-member(1))|如果集合仅包含一个项，则此方法返回该项。|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-items-member)|获取此集合中已加载的子项。|
|[Document](/javascript/api/word/word.document)|[customXmlParts](/javascript/api/word/word.document#word-word-document-customxmlparts-member)|获取文档中的自定义 XML 部件。|
||[deleteBookmark (name： string) ](/javascript/api/word/word.document#word-word-document-deletebookmark-member(1))|从文档中删除书签（如果存在）。|
||[getBookmarkRange (名称：string) ](/javascript/api/word/word.document#word-word-document-getbookmarkrange-member(1))|获取书签的范围。|
||[getBookmarkRangeOrNullObject (name： string) ](/javascript/api/word/word.document#word-word-document-getbookmarkrangeornullobject-member(1))|获取书签的范围。|
||[ignorePunct](/javascript/api/word/word.document#word-word-document-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.document#word-word-document-ignorespace-member)||
||[matchCase](/javascript/api/word/word.document#word-word-document-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.document#word-word-document-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.document#word-word-document-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.document#word-word-document-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.document#word-word-document-matchwildcards-member)||
||[onContentControlAdded](/javascript/api/word/word.document#word-word-document-oncontentcontroladded-member)|添加内容控件时发生。|
||[search (searchText： string， searchOptions？： Word.SearchOptions \| { ignorePunct？： boolean ignoreSpace？： boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： boolean matchWholeWord？： boolean matchWildcards？： boolean }) ](/javascript/api/word/word.document#word-word-document-search-member(1))|使用指定的搜索选项搜索整个文档的范围。|
||[设置](/javascript/api/word/word.document#word-word-document-settings-member)|获取文档中加载项的设置。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[customXmlParts](/javascript/api/word/word.documentcreated#word-word-documentcreated-customxmlparts-member)|获取文档中的自定义 XML 部件。|
||[deleteBookmark (name： string) ](/javascript/api/word/word.documentcreated#word-word-documentcreated-deletebookmark-member(1))|从文档中删除书签（如果存在）。|
||[getBookmarkRange (name： string) ](/javascript/api/word/word.documentcreated#word-word-documentcreated-getbookmarkrange-member(1))|获取书签的范围。|
||[getBookmarkRangeOrNullObject (name： string) ](/javascript/api/word/word.documentcreated#word-word-documentcreated-getbookmarkrangeornullobject-member(1))|获取书签的范围。|
||[设置](/javascript/api/word/word.documentcreated#word-word-documentcreated-settings-member)|获取文档中加载项的设置。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-imageformat-member)|获取内嵌图像的格式。|
|[列表](/javascript/api/word/word.list)|[getLevelFont (级别：number) ](/javascript/api/word/word.list#word-word-list-getlevelfont-member(1))|获取列表中指定级别的项目符号、编号或图片的字体。|
||[getLevelPicture (级别：number) ](/javascript/api/word/word.list#word-word-list-getlevelpicture-member(1))|获取列表中指定级别的图片的 base64 编码字符串表示形式。|
||[resetLevelFont (level： number， resetFontName？： boolean) ](/javascript/api/word/word.list#word-word-list-resetlevelfont-member(1))|重置列表中指定级别的项目符号、编号或图片的字体。|
||[setLevelPicture (级别： number， base64EncodedImage？： string) ](/javascript/api/word/word.list#word-word-list-setlevelpicture-member(1))|设置列表中指定级别的图片。|
|[区域](/javascript/api/word/word.range)|[getBookmarks (includeHidden？： boolean， includeAdjacent？： boolean) ](/javascript/api/word/word.range#word-word-range-getbookmarks-member(1))|获取区域内的所有书签或与区域重叠的名称。|
||[insertBookmark (name： string) ](/javascript/api/word/word.range#word-word-range-insertbookmark-member(1))|在范围中插入书签。|
|[设置](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#word-word-setting-delete-member(1))|删除 Setting 对象。|
||[key](/javascript/api/word/word.setting#word-word-setting-key-member)|获取设置的键。|
||[value](/javascript/api/word/word.setting#word-word-setting-value-member)|获取或设置设置的值。|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add (key： string， value： any) ](/javascript/api/word/word.settingcollection#word-word-settingcollection-add-member(1))|创建新设置或设置现有设置。|
||[deleteAll () ](/javascript/api/word/word.settingcollection#word-word-settingcollection-deleteall-member(1))|删除此外接程序中的所有设置。|
||[getCount()](/javascript/api/word/word.settingcollection#word-word-settingcollection-getcount-member(1))|获取设置计数。|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getitem-member(1))|按其键（区分大小写）获取 setting 对象。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getitemornullobject-member(1))|按其键（区分大小写）获取 setting 对象。|
||[items](/javascript/api/word/word.settingcollection#word-word-settingcollection-items-member)|获取此集合中已加载的子项。|
|[Table](/javascript/api/word/word.table)|[mergeCells (topRow： number， firstCell： number， bottomRow： number， lastCell： number) ](/javascript/api/word/word.table#word-word-table-mergecells-member(1))|合并第一个单元格和最后一个单元格（包含边界）的单元格。|
|[TableCell](/javascript/api/word/word.tablecell)|[split (rowCount： number， columnCount： number) ](/javascript/api/word/word.tablecell#word-word-tablecell-split-member(1))|将单元格拆分为指定的行数和列数。|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#word-word-tablerow-insertcontentcontrol-member(1))|在行上插入内容控件。|
||[merge () ](/javascript/api/word/word.tablerow#word-word-tablerow-merge-member(1))|将行合并为一个单元格。|

## <a name="web-only-api-list"></a>仅 Web API 列表

下表列出了当前仅在 Word web 版 预览版中的 Word JavaScript API。 若要查看所有 Word JavaScript API 的完整列表， (预览 API 和以前发布的 API) ，请参阅 [所有 Word JavaScript API](/javascript/api/word?view=word-js-preview&preserve-view=true)。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[正文](/javascript/api/word/word.body)|[endnotes](/javascript/api/word/word.body#word-word-body-endnotes-member)|获取正文中的尾注集合。|
||[脚注](/javascript/api/word/word.body#word-word-body-footnotes-member)|获取正文中的脚注集合。|
||[getComments () ](/javascript/api/word/word.body#word-word-body-getcomments-member(1))|获取与正文关联的注释。|
||[getReviewedText (changeTrackingVersion？：Word.ChangeTrackingVersion) ](/javascript/api/word/word.body#word-word-body-getreviewedtext-member(1))|根据 ChangeTrackingVersion 选择获取已审阅文本。|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|获取 body 的类型。|
|[Comment](/javascript/api/word/word.comment)|[authorEmail](/javascript/api/word/word.comment#word-word-comment-authoremail-member)|获取批注作者的电子邮件。|
||[authorName](/javascript/api/word/word.comment#word-word-comment-authorname-member)|获取批注作者的姓名。|
||[content](/javascript/api/word/word.comment#word-word-comment-content-member)|获取或设置批注的内容为纯文本。|
||[contentRange](/javascript/api/word/word.comment#word-word-comment-contentrange-member)|获取或设置注释线程状态。|
||[creationDate](/javascript/api/word/word.comment#word-word-comment-creationdate-member)|获取批注的创建日期。|
||[delete()](/javascript/api/word/word.comment#word-word-comment-delete-member(1))|删除注释及其回复。|
||[getRange()](/javascript/api/word/word.comment#word-word-comment-getrange-member(1))|获取批注位于主文档中的范围。|
||[id](/javascript/api/word/word.comment#word-word-comment-id-member)|ID|
||[replies](/javascript/api/word/word.comment#word-word-comment-replies-member)|获取与注释关联的 reply 对象的集合。|
||[reply (replyText： string) ](/javascript/api/word/word.comment#word-word-comment-reply-member(1))|将新回复添加到注释线程的末尾。|
||[已解决](/javascript/api/word/word.comment#word-word-comment-resolved-member)|获取或设置注释线程的状态。|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[getFirst()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirst-member(1))|获取集合中的第一个注释。|
||[getFirstOrNullObject () ](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirstornullobject-member(1))|获取集合中的第一个注释。|
||[getItem(index: number)](/javascript/api/word/word.commentcollection#word-word-commentcollection-getitem-member(1))|按注释对象在集合中的索引获取该对象。|
||[items](/javascript/api/word/word.commentcollection#word-word-commentcollection-items-member)|获取此集合中已加载的子项。|
|[CommentContentRange](/javascript/api/word/word.commentcontentrange)|[bold](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-bold-member)|获取或设置一个值，该值指示批注文本是否加粗。|
||[hyperlink](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-hyperlink-member)|获取 range 内的第一个超链接，或在 range 内设置超链接。|
||[insertText (text： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-inserttext-member(1))|将文本插入到指定位置。|
||[isEmpty](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-isempty-member)|检查 range 长度是否为零。|
||[italic](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-italic-member)|获取或设置一个值，该值指示批注文本是否为 italicized。|
||[strikeThrough](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-strikethrough-member)|获取或设置一个值，该值指示批注文本是否有删除线。|
||[text](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-text-member)|获取批注区域的文本。|
||[underline](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-underline-member)|获取或设置一个值，该值指示批注文本的下划线类型。|
|[CommentReply](/javascript/api/word/word.commentreply)|[authorEmail](/javascript/api/word/word.commentreply#word-word-commentreply-authoremail-member)|获取批注回复作者的电子邮件。|
||[authorName](/javascript/api/word/word.commentreply#word-word-commentreply-authorname-member)|获取批注回复作者的姓名。|
||[content](/javascript/api/word/word.commentreply#word-word-commentreply-content-member)|获取或设置批注回复的内容。|
||[contentRange](/javascript/api/word/word.commentreply#word-word-commentreply-contentrange-member)|获取或设置 commentReply 的内容范围。|
||[creationDate](/javascript/api/word/word.commentreply#word-word-commentreply-creationdate-member)|获取批注回复的创建日期。|
||[delete()](/javascript/api/word/word.commentreply#word-word-commentreply-delete-member(1))|删除批注回复。|
||[id](/javascript/api/word/word.commentreply#word-word-commentreply-id-member)|ID|
||[parentComment](/javascript/api/word/word.commentreply#word-word-commentreply-parentcomment-member)|获取此回复的父批注。|
|[CommentReplyCollection](/javascript/api/word/word.commentreplycollection)|[getFirst()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirst-member(1))|获取集合中的第一个批注回复。|
||[getFirstOrNullObject () ](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirstornullobject-member(1))|获取集合中的第一个批注回复。|
||[getItem(index: number)](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getitem-member(1))|按注释答复对象在集合中的索引获取该对象。|
||[items](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-items-member)|获取此集合中已加载的子项。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[endnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-endnotes-member)|获取 contentcontrol 中的尾注集合。|
||[脚注](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-footnotes-member)|获取 contentcontrol 中的脚注集合。|
||[getComments () ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getcomments-member(1))|获取与正文关联的注释。|
||[getReviewedText (changeTrackingVersion？：Word.ChangeTrackingVersion) ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getreviewedtext-member(1))|根据 ChangeTrackingVersion 选择获取已审阅文本。|
|[Document](/javascript/api/word/word.document)|[changeTrackingMode](/javascript/api/word/word.document#word-word-document-changetrackingmode-member)|获取或设置 ChangeTracking 模式。|
||[getEndnoteBody () ](/javascript/api/word/word.document#word-word-document-getendnotebody-member(1))|获取单个正文中的文档的尾注。|
||[getFootnoteBody () ](/javascript/api/word/word.document#word-word-document-getfootnotebody-member(1))|获取单个正文中的文档脚注。|
|[NoteItem](/javascript/api/word/word.noteitem)|[body](/javascript/api/word/word.noteitem#word-word-noteitem-body-member)|表示便笺项目的 body 对象。|
||[delete()](/javascript/api/word/word.noteitem#word-word-noteitem-delete-member(1))|删除便笺项目。|
||[getNext () ](/javascript/api/word/word.noteitem#word-word-noteitem-getnext-member(1))|获取同一类型的下一个便笺项。|
||[getNextOrNullObject () ](/javascript/api/word/word.noteitem#word-word-noteitem-getnextornullobject-member(1))|获取同一类型的下一个便笺项。|
||[reference](/javascript/api/word/word.noteitem#word-word-noteitem-reference-member)|代表主文档中的脚注或尾注引用。|
||[type](/javascript/api/word/word.noteitem#word-word-noteitem-type-member)|代表便笺项目类型：脚注或尾注。|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirst-member(1))|获取此集合中的第一个便笺项。|
||[getFirstOrNullObject () ](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirstornullobject-member(1))|获取此集合中的第一个便笺项。|
||[items](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-items-member)|获取此集合中已加载的子项。|
|[Paragraph](/javascript/api/word/word.paragraph)|[endnotes](/javascript/api/word/word.paragraph#word-word-paragraph-endnotes-member)|获取段落中的尾注集合。|
||[脚注](/javascript/api/word/word.paragraph#word-word-paragraph-footnotes-member)|获取段落中的脚注集合。|
||[getComments () ](/javascript/api/word/word.paragraph#word-word-paragraph-getcomments-member(1))|获取与段落关联的注释。|
||[getReviewedText (changeTrackingVersion？：Word.ChangeTrackingVersion) ](/javascript/api/word/word.paragraph#word-word-paragraph-getreviewedtext-member(1))|根据 ChangeTrackingVersion 选择获取已审阅文本。|
|[区域](/javascript/api/word/word.range)|[endnotes](/javascript/api/word/word.range#word-word-range-endnotes-member)|获取范围中的尾注集合。|
||[脚注](/javascript/api/word/word.range#word-word-range-footnotes-member)|获取范围中的脚注集合。|
||[getComments () ](/javascript/api/word/word.range#word-word-range-getcomments-member(1))|获取与区域关联的注释。|
||[getReviewedText (changeTrackingVersion？：Word.ChangeTrackingVersion) ](/javascript/api/word/word.range#word-word-range-getreviewedtext-member(1))|根据 ChangeTrackingVersion 选择获取已审阅文本。|
||[insertComment (commentText： string) ](/javascript/api/word/word.range#word-word-range-insertcomment-member(1))|在范围中插入注释。|
||[insertEndnote (insertText？： string) ](/javascript/api/word/word.range#word-word-range-insertendnote-member(1))|插入尾注。|
||[insertFootnote (insertText？： string) ](/javascript/api/word/word.range#word-word-range-insertfootnote-member(1))|插入脚注。|
|[Table](/javascript/api/word/word.table)|[endnotes](/javascript/api/word/word.table#word-word-table-endnotes-member)|获取 table 中的尾注集合。|
||[脚注](/javascript/api/word/word.table#word-word-table-footnotes-member)|获取表格中脚注的集合。|
|[TableRow](/javascript/api/word/word.tablerow)|[endnotes](/javascript/api/word/word.tablerow#word-word-tablerow-endnotes-member)|获取表格行中的尾注集合。|
||[脚注](/javascript/api/word/word.tablerow#word-word-tablerow-footnotes-member)|获取表格行中的脚注集合。|

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Word JavaScript API 要求集](word-api-requirement-sets.md)

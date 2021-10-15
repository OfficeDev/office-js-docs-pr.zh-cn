---
title: Word JavaScript API 要求集 1.3
description: 有关 WordApi 1.3 要求集的详细信息。
ms.date: 03/09/2021
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: 3d569390a87b3eb153f8139c9e9608bf747fde07
ms.sourcegitcommit: 3b187769e86530334ca83cfdb03c1ecfac2ad9a8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/15/2021
ms.locfileid: "60367465"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Word JavaScript API 1.3 的最近更新

WordApi 1.3 增加了对内容控件和文档级别设置的更多支持。

## <a name="api-list"></a>API 列表

下表列出了 Word JavaScript API 要求集 1.3 中的 API。 若要查看受 Word JavaScript API 要求集 1.3 或更早版本支持的所有 API 的 API 参考文档，请参阅要求集 [1.3](/javascript/api/word?view=word-js-1.3&preserve-view=true)或更早中的 Word API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument (base64File？： string) ](/javascript/api/word/word.application#createDocument_base64File_)|使用可选的 base64 编码文档文件创建新.docx文档。|
|[正文](/javascript/api/word/word.body)|[getRange (rangeLocation？：Word.RangeLocation) ](/javascript/api/word/word.body#getRange_rangeLocation_)|获取整个正文或正文的起点/终点，作为一个范围。|
||[insertTable (rowCount： number， columnCount： number， insertLocation： Word.InsertLocation， values？： string[][]) ](/javascript/api/word/word.body#insertTable_rowCount__columnCount__insertLocation__values_)|插入包含指定行数和列数的 table。|
||[lists](/javascript/api/word/word.body#lists)|获取 body 中的一组 list 对象。|
||[parentBody](/javascript/api/word/word.body#parentBody)|获取 body 的父正文。|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentBodyOrNullObject)|获取 body 的父正文。|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentContentControlOrNullObject)|获取包含正文的内容控件。|
||[parentSection](/javascript/api/word/word.body#parentSection)|获取 body 的父节。|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentSectionOrNullObject)|获取 body 的父节。|
||[styleBuiltIn](/javascript/api/word/word.body#styleBuiltIn)|获取或设置 body 的嵌入样式名称。|
||[表](/javascript/api/word/word.body#tables)|获取 body 中的一组 table 对象。|
||[type](/javascript/api/word/word.body#type)|获取 body 的类型。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange (rangeLocation？：Word.RangeLocation) ](/javascript/api/word/word.contentcontrol#getRange_rangeLocation_)|获取整个内容控件或内容控件的起点/终点，作为一个范围。|
||[getTextRanges (结束Marks： string[]， trimSpacing？： boolean) ](/javascript/api/word/word.contentcontrol#getTextRanges_endingMarks__trimSpacing_)|使用标点符号和/或其他结束标记获取内容控件中的文本范围。|
||[insertTable (rowCount： number， columnCount： number， insertLocation： Word.InsertLocation， values？： string[][]) ](/javascript/api/word/word.contentcontrol#insertTable_rowCount__columnCount__insertLocation__values_)|将包含指定行数和列数的 table 插入 contentControl 中或在其旁边插入。|
||[lists](/javascript/api/word/word.contentcontrol#lists)|获取 contentControl 中的一组 list 对象。|
||[parentBody](/javascript/api/word/word.contentcontrol#parentBody)|获取 contentControl 的父正文。|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentContentControlOrNullObject)|获取包含此内容控件的内容控件。|
||[parentTable](/javascript/api/word/word.contentcontrol#parentTable)|获取包含 contentControl 的 table。|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parentTableCell)|获取包含 contentControl 的 tableCell。|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parentTableCellOrNullObject)|获取包含 contentControl 的 tableCell。|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parentTableOrNullObject)|获取包含 contentControl 的 table。|
||[split (delimiters： string[]， multiParagraphs？： boolean， trimDelimiters？： boolean， trimSpacing？： boolean) ](/javascript/api/word/word.contentcontrol#split_delimiters__multiParagraphs__trimDelimiters__trimSpacing_)|使用分隔符将内容控件拆分为各个子范围。|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#styleBuiltIn)|获取或设置 contentControl 的嵌入样式名称。|
||[subtype](/javascript/api/word/word.contentcontrol#subtype)|获取 contentControl 的子类型。|
||[表](/javascript/api/word/word.contentcontrol#tables)|获取 contentControl 中的一组 table 对象。|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject (id： number) ](/javascript/api/word/word.contentcontrolcollection#getByIdOrNullObject_id_)|按其标识符获取内容控件。|
||[getByTypes (类型：Word.ContentControlType[]) ](/javascript/api/word/word.contentcontrolcollection#getByTypes_types_)|获取具有指定类型和/或子类型的内容控件。|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getFirst__)|获取此集合中的第一个内容控件。|
||[getFirstOrNullObject () ](/javascript/api/word/word.contentcontrolcollection#getFirstOrNullObject__)|获取此集合中的第一个内容控件。|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete__)|删除 custom property 对象。|
||[key](/javascript/api/word/word.customproperty#key)|获取 customProperty 的键。|
||[type](/javascript/api/word/word.customproperty#type)|获取自定义属性的值类型。|
||[value](/javascript/api/word/word.customproperty#value)|获取或设置自定义属性的值。|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add (key： string， value： any) ](/javascript/api/word/word.custompropertycollection#add_key__value_)|新建自定义属性或设置现有自定义属性。|
||[deleteAll () ](/javascript/api/word/word.custompropertycollection#deleteAll__)|删除此集合中的所有自定义属性。|
||[getCount()](/javascript/api/word/word.custompropertycollection#getCount__)|获取自定义属性的计数。|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getItem_key_)|按键获取自定义属性对象（不区分大小写）。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getItemOrNullObject_key_)|按键获取自定义属性对象（不区分大小写）。|
||[items](/javascript/api/word/word.custompropertycollection#items)|获取此集合中已加载的子项。|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|获取文档的属性。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[body](/javascript/api/word/word.documentcreated#body)|获取文档的 body 对象。|
||[contentControls](/javascript/api/word/word.documentcreated#contentControls)|获取文档中的内容控件对象的集合。|
||[打开 () ](/javascript/api/word/word.documentcreated#open__)|打开文档。|
||[properties](/javascript/api/word/word.documentcreated#properties)|获取文档的属性。|
||[save()](/javascript/api/word/word.documentcreated#save__)|保存文档。|
||[saved](/javascript/api/word/word.documentcreated#saved)|指示是否已保存在文档中所做的更改。|
||[sections](/javascript/api/word/word.documentcreated#sections)|获取文档中 section 对象的集合。|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[applicationName](/javascript/api/word/word.documentproperties#applicationName)|获取 document 的应用程序名称。|
||[author](/javascript/api/word/word.documentproperties#author)|获取或设置 document 的作者。|
||[category](/javascript/api/word/word.documentproperties#category)|获取或设置 document 的类别。|
||[comments](/javascript/api/word/word.documentproperties#comments)|获取或设置 document 的注释。|
||[company](/javascript/api/word/word.documentproperties#company)|获取或设置 document 的公司。|
||[creationDate](/javascript/api/word/word.documentproperties#creationDate)|获取文档的创建日期。|
||[customProperties](/javascript/api/word/word.documentproperties#customProperties)|获取 document 的一组 customProperty。|
||[format](/javascript/api/word/word.documentproperties#format)|获取或设置 document 的格式。|
||[keywords](/javascript/api/word/word.documentproperties#keywords)|获取或设置 document 的关键字。|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastAuthor)|获取文档的最后一个作者。|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastPrintDate)|获取文档的上次打印日期。|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastSaveTime)|获取 document 的上次保存日期。|
||[manager](/javascript/api/word/word.documentproperties#manager)|获取或设置 document 的管理者。|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionNumber)|获取文档的修订号。|
||[security](/javascript/api/word/word.documentproperties#security)|获取文档的安全设置。|
||[subject](/javascript/api/word/word.documentproperties#subject)|获取或设置 document 的主题。|
||[template](/javascript/api/word/word.documentproperties#template)|获取文档的模板。|
||[title](/javascript/api/word/word.documentproperties#title)|获取或设置 document 的标题。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext () ](/javascript/api/word/word.inlinepicture#getNext__)|获取下一个嵌入式图像。|
||[getNextOrNullObject () ](/javascript/api/word/word.inlinepicture#getNextOrNullObject__)|获取下一个嵌入式图像。|
||[getRange (rangeLocation？：Word.RangeLocation) ](/javascript/api/word/word.inlinepicture#getRange_rangeLocation_)|获取图片或图片的起点/终点，作为一个范围。|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentContentControlOrNullObject)|获取包含嵌入式图像的内容控件。|
||[parentTable](/javascript/api/word/word.inlinepicture#parentTable)|获取包含嵌入式图像的 table。|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parentTableCell)|获取包含嵌入式图像的 tableCell。|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parentTableCellOrNullObject)|获取包含嵌入式图像的 tableCell。|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parentTableOrNullObject)|获取包含嵌入式图像的 table。|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getFirst__)|获取此集合中的第一个嵌入式图像。|
||[getFirstOrNullObject () ](/javascript/api/word/word.inlinepicturecollection#getFirstOrNullObject__)|获取此集合中的第一个嵌入式图像。|
|[列表](/javascript/api/word/word.list)|[getLevelParagraphs (级别：number) ](/javascript/api/word/word.list#getLevelParagraphs_level_)|获取列表中指定级别的段落。|
||[getLevelString (级别：number) ](/javascript/api/word/word.list#getLevelString_level_)|以字符串形式获取指定级别的项目符号、编号或图片。|
||[id](/javascript/api/word/word.list#id)|获取列表的 ID。|
||[insertParagraph (paragraphText： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.list#insertParagraph_paragraphText__insertLocation_)|在指定位置插入段落。|
||[levelExistences](/javascript/api/word/word.list#levelExistences)|检查 list 中是否包含所有 9 个级别。|
||[levelTypes](/javascript/api/word/word.list#levelTypes)|获取 list 中的所有 9 个级别类型。|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|获取 list 中的段落。|
||[setLevelAlignment (level： number， alignment： Word.Alignment) ](/javascript/api/word/word.list#setLevelAlignment_level__alignment_)|设置项目符号、编号或图片在列表中指定级别的对齐方式。|
||[setLevelBullet (level： number， listBullet： Word.ListBullet， charCode？： number， fontName？： string) ](/javascript/api/word/word.list#setLevelBullet_level__listBullet__charCode__fontName_)|设置 list 中指定级别的项目符号格式。|
||[setLevelIndents (level： number， textIndent： number， bulletNumberPictureIndent： number) ](/javascript/api/word/word.list#setLevelIndents_level__textIndent__bulletNumberPictureIndent_)|设置列表中指定级别的两种缩进方式。|
||[setLevelNumbering (level： number， listNumbering： Word.ListNumbering， formatString？： Array<string \| number>) ](/javascript/api/word/word.list#setLevelNumbering_level__listNumbering__formatString_)|设置列表中指定级别的编号格式。|
||[setLevelStartingNumber (level： number， startingNumber： number) ](/javascript/api/word/word.list#setLevelStartingNumber_level__startingNumber_)|设置 list 中指定级别的起始编号。|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getById_id_)|按标识符获取列表。|
||[getByIdOrNullObject (id： number) ](/javascript/api/word/word.listcollection#getByIdOrNullObject_id_)|按标识符获取列表。|
||[getFirst()](/javascript/api/word/word.listcollection#getFirst__)|获取此集合中的第一个列表。|
||[getFirstOrNullObject () ](/javascript/api/word/word.listcollection#getFirstOrNullObject__)|获取此集合中的第一个列表。|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getItem_index_)|按列表对象在集合中的索引获取列表。|
||[items](/javascript/api/word/word.listcollection#items)|获取此集合中已加载的子项。|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor (parentOnly？： boolean) ](/javascript/api/word/word.listitem#getAncestor_parentOnly_)|获取列表项目的父级或最近的上级元素（如果父级不存在的话）。|
||[getAncestorOrNullObject (parentOnly？： boolean) ](/javascript/api/word/word.listitem#getAncestorOrNullObject_parentOnly_)|获取列表项目的父级或最近的上级元素（如果父级不存在的话）。|
||[getDescendants (directChildrenOnly？： boolean) ](/javascript/api/word/word.listitem#getDescendants_directChildrenOnly_)|获取相应列表项目的所有后代列表项目。|
||[level](/javascript/api/word/word.listitem#level)|获取或设置 list 中项的级别。|
||[listString](/javascript/api/word/word.listitem#listString)|获取字符串形式的列表项项目符号、编号或图片。|
||[siblingIndex](/javascript/api/word/word.listitem#siblingIndex)|获取 listItem 相对于同级元素的序号。|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList (listId： number， level： number) ](/javascript/api/word/word.paragraph#attachToList_listId__level_)|将 paragraph 加入指定级别的现有 list。|
||[detachFromList () ](/javascript/api/word/word.paragraph#detachFromList__)|如果此段落是列表项目的话，从列表中移出此段落。|
||[getNext () ](/javascript/api/word/word.paragraph#getNext__)|获取下一个段落。|
||[getNextOrNullObject () ](/javascript/api/word/word.paragraph#getNextOrNullObject__)|获取下一个段落。|
||[getPrevious () ](/javascript/api/word/word.paragraph#getPrevious__)|获取上一个段落。|
||[getPreviousOrNullObject () ](/javascript/api/word/word.paragraph#getPreviousOrNullObject__)|获取上一个段落。|
||[getRange (rangeLocation？：Word.RangeLocation) ](/javascript/api/word/word.paragraph#getRange_rangeLocation_)|获取整个段落或段落的起点/终点，作为一个范围。|
||[getTextRanges (结束Marks： string[]， trimSpacing？： boolean) ](/javascript/api/word/word.paragraph#getTextRanges_endingMarks__trimSpacing_)|使用标点符号和/或其他结束标记获取段落中的文本范围。|
||[insertTable (rowCount： number， columnCount： number， insertLocation： Word.InsertLocation， values？： string[][]) ](/javascript/api/word/word.paragraph#insertTable_rowCount__columnCount__insertLocation__values_)|插入包含指定行数和列数的 table。|
||[isLastParagraph](/javascript/api/word/word.paragraph#isLastParagraph)|指明 paragraph 是其父正文内的最后一个段落。|
||[isListItem](/javascript/api/word/word.paragraph#isListItem)|检查 paragraph 是否为 listItem。|
||[list](/javascript/api/word/word.paragraph#list)|获取 paragraph 所属的 List。|
||[listItem](/javascript/api/word/word.paragraph#listItem)|获取 paragraph 的 ListItem。|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listItemOrNullObject)|获取 paragraph 的 ListItem。|
||[listOrNullObject](/javascript/api/word/word.paragraph#listOrNullObject)|获取 paragraph 所属的 List。|
||[parentBody](/javascript/api/word/word.paragraph#parentBody)|获取 paragraph 的父正文。|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentContentControlOrNullObject)|获取包含段落的内容控件。|
||[parentTable](/javascript/api/word/word.paragraph#parentTable)|获取包含 paragraph 的 table。|
||[parentTableCell](/javascript/api/word/word.paragraph#parentTableCell)|获取包含 paragraph 的 tableCell。|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parentTableCellOrNullObject)|获取包含 paragraph 的 tableCell。|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parentTableOrNullObject)|获取包含 paragraph 的 table。|
||[split (delimiters： string[]， trimDelimiters？： boolean， trimSpacing？： boolean) ](/javascript/api/word/word.paragraph#split_delimiters__trimDelimiters__trimSpacing_)|使用分隔符将段落拆分为多个子范围。|
||[startNewList () ](/javascript/api/word/word.paragraph#startNewList__)|生成包含此 paragraph 的新 list。|
||[styleBuiltIn](/javascript/api/word/word.paragraph#styleBuiltIn)|获取或设置 paragraph 的嵌入样式名称。|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tableNestingLevel)|获取 paragraph 的表级别。|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getFirst__)|获取此集合中的第一个段落。|
||[getFirstOrNullObject () ](/javascript/api/word/word.paragraphcollection#getFirstOrNullObject__)|获取此集合中的第一个段落。|
||[getLast () ](/javascript/api/word/word.paragraphcollection#getLast__)|获取此集合中的最后一个段落。|
||[getLastOrNullObject () ](/javascript/api/word/word.paragraphcollection#getLastOrNullObject__)|获取此集合中的最后一个段落。|
|[区域](/javascript/api/word/word.range)|[compareLocationWith (range： Word.Range) ](/javascript/api/word/word.range#compareLocationWith_range_)|比较此范围与另一范围的位置。|
||[expandTo (range： Word.Range) ](/javascript/api/word/word.range#expandTo_range_)|返回从此 range 进行任一方向扩展的新 range，以便覆盖另一 range。|
||[expandToOrNullObject (范围：Word.Range) ](/javascript/api/word/word.range#expandToOrNullObject_range_)|返回从此 range 进行任一方向扩展的新 range，以便覆盖另一 range。|
||[getHyperlinkRanges () ](/javascript/api/word/word.range#getHyperlinkRanges__)|获取相应范围内的超链接子范围。|
||[getNextTextRange (结束Marks： string[]， trimSpacing？： boolean) ](/javascript/api/word/word.range#getNextTextRange_endingMarks__trimSpacing_)|使用标点符号和/或其他结束标记获取下一个文本范围。|
||[getNextTextRangeOrNullObject (结束Marks： string[]， trimSpacing？： boolean) ](/javascript/api/word/word.range#getNextTextRangeOrNullObject_endingMarks__trimSpacing_)|使用标点符号和/或其他结束标记获取下一个文本范围。|
||[getRange (rangeLocation？：Word.RangeLocation) ](/javascript/api/word/word.range#getRange_rangeLocation_)|克隆相应范围，或获取该范围的起点/终点作为一个新范围。|
||[getTextRanges (结束Marks： string[]， trimSpacing？： boolean) ](/javascript/api/word/word.range#getTextRanges_endingMarks__trimSpacing_)|使用标点符号和/或其他结束标记获取范围中的文本子范围。|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|获取 range 内的第一个超链接，或在 range 内设置超链接。|
||[insertTable (rowCount： number， columnCount： number， insertLocation： Word.InsertLocation， values？： string[][]) ](/javascript/api/word/word.range#insertTable_rowCount__columnCount__insertLocation__values_)|插入包含指定行数和列数的 table。|
||[intersectWith (range： Word.Range) ](/javascript/api/word/word.range#intersectWith_range_)|返回新 range 作为此 range 与另一 range 的交集。|
||[intersectWithOrNullObject (范围：Word.Range) ](/javascript/api/word/word.range#intersectWithOrNullObject_range_)|返回新 range 作为此 range 与另一 range 的交集。|
||[isEmpty](/javascript/api/word/word.range#isEmpty)|检查 range 长度是否为零。|
||[lists](/javascript/api/word/word.range#lists)|获取 range 中的一组 list 对象。|
||[parentBody](/javascript/api/word/word.range#parentBody)|获取 range 的父正文。|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentContentControlOrNullObject)|获取包含该范围的内容控件。|
||[parentTable](/javascript/api/word/word.range#parentTable)|获取包含 range 的 table。|
||[parentTableCell](/javascript/api/word/word.range#parentTableCell)|获取包含 range 的 tableCell。|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parentTableCellOrNullObject)|获取包含 range 的 tableCell。|
||[parentTableOrNullObject](/javascript/api/word/word.range#parentTableOrNullObject)|获取包含 range 的 table。|
||[split (delimiters： string[]， multiParagraphs？： boolean， trimDelimiters？： boolean， trimSpacing？： boolean) ](/javascript/api/word/word.range#split_delimiters__multiParagraphs__trimDelimiters__trimSpacing_)|使用分隔符将相应范围拆分为各个子范围。|
||[styleBuiltIn](/javascript/api/word/word.range#styleBuiltIn)|获取或设置 range 的嵌入样式名称。|
||[表](/javascript/api/word/word.range#tables)|获取 range 中的一组 table 对象。|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getFirst__)|获取此集合中的第一个范围。|
||[getFirstOrNullObject () ](/javascript/api/word/word.rangecollection#getFirstOrNullObject__)|获取此集合中的第一个范围。|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#application)|[Api 集：WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext () ](/javascript/api/word/word.section#getNext__)|获取下一节。|
||[getNextOrNullObject () ](/javascript/api/word/word.section#getNextOrNullObject__)|获取下一节。|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getFirst__)|获取此集合中的第一节。|
||[getFirstOrNullObject () ](/javascript/api/word/word.sectioncollection#getFirstOrNullObject__)|获取此集合中的第一节。|
|[Table](/javascript/api/word/word.table)|[addColumns (insertLocation： Word.InsertLocation， columnCount： number， values？： string[][]) ](/javascript/api/word/word.table#addColumns_insertLocation__columnCount__values_)|使用第一个或最后一个现有列作为模板，将列添加到 table 的开头或结尾。|
||[addRows (insertLocation： Word.InsertLocation， rowCount： number， values？： string[][]) ](/javascript/api/word/word.table#addRows_insertLocation__rowCount__values_)|使用第一个或最后一个现有行作为模板，将行添加到 table 的开头或结尾。|
||[alignment](/javascript/api/word/word.table#alignment)|获取或设置表格与页面列的对齐方式。|
||[autoFitWindow () ](/javascript/api/word/word.table#autoFitWindow__)|自动调整表列，以适应窗口的宽度。|
||[clear()](/javascript/api/word/word.table#clear__)|清除表内容。|
||[delete()](/javascript/api/word/word.table#delete__)|删除整个表格。|
||[deleteColumns (columnIndex： number， columnCount？： number) ](/javascript/api/word/word.table#deleteColumns_columnIndex__columnCount_)|删除特定的列。|
||[deleteRows (rowIndex： number， rowCount？： number) ](/javascript/api/word/word.table#deleteRows_rowIndex__rowCount_)|删除特定的行。|
||[distributeColumns () ](/javascript/api/word/word.table#distributeColumns__)|将列设置为等宽。|
||[font](/javascript/api/word/word.table#font)|获取字体。|
||[getBorder (borderLocation：Word.BorderLocation) ](/javascript/api/word/word.table#getBorder_borderLocation_)|获取指定边框的边框样式。|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getCell_rowIndex__cellIndex_)|获取指定行和列处的表单元格。|
||[getCellOrNullObject (rowIndex： number， cellIndex： number) ](/javascript/api/word/word.table#getCellOrNullObject_rowIndex__cellIndex_)|获取指定行和列处的表单元格。|
||[getCellPadding (cellPaddingLocation：Word.CellPaddingLocation) ](/javascript/api/word/word.table#getCellPadding_cellPaddingLocation_)|获取单元格填充（以磅为单位）。|
||[getNext () ](/javascript/api/word/word.table#getNext__)|获取下一个表格。|
||[getNextOrNullObject () ](/javascript/api/word/word.table#getNextOrNullObject__)|获取下一个表格。|
||[getParagraphAfter () ](/javascript/api/word/word.table#getParagraphAfter__)|获取 table 之后的 paragraph。|
||[getParagraphAfterOrNullObject () ](/javascript/api/word/word.table#getParagraphAfterOrNullObject__)|获取 table 之后的 paragraph。|
||[getParagraphBefore () ](/javascript/api/word/word.table#getParagraphBefore__)|获取 table 之前的 paragraph。|
||[getParagraphBeforeOrNullObject () ](/javascript/api/word/word.table#getParagraphBeforeOrNullObject__)|获取 table 之前的 paragraph。|
||[getRange (rangeLocation？：Word.RangeLocation) ](/javascript/api/word/word.table#getRange_rangeLocation_)|获取包含此表格的范围，或包含此表格的开头或结尾的范围。|
||[headerRowCount](/javascript/api/word/word.table#headerRowCount)|获取并设置标题行数。|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalAlignment)|获取并设置 table 中每个单元格的水平对齐方式。|
||[ignorePunct](/javascript/api/word/word.table#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.table#ignoreSpace)||
||[insertContentControl()](/javascript/api/word/word.table#insertContentControl__)|在表格中插入内容控件。|
||[insertParagraph (paragraphText： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.table#insertParagraph_paragraphText__insertLocation_)|在指定位置插入段落。|
||[insertTable (rowCount： number， columnCount： number， insertLocation： Word.InsertLocation， values？： string[][]) ](/javascript/api/word/word.table#insertTable_rowCount__columnCount__insertLocation__values_)|插入包含指定行数和列数的 table。|
||[isUniform](/javascript/api/word/word.table#isUniform)|指明所有表行是否一致。|
||[matchCase](/javascript/api/word/word.table#matchCase)||
||[matchPrefix](/javascript/api/word/word.table#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.table#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.table#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.table#matchWildcards)||
||[nestingLevel](/javascript/api/word/word.table#nestingLevel)|获取 table 的嵌套级别。|
||[parentBody](/javascript/api/word/word.table#parentBody)|获取 table 的父正文。|
||[parentContentControl](/javascript/api/word/word.table#parentContentControl)|获取包含 table 的 contentControl。|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentContentControlOrNullObject)|获取包含 table 的 contentControl。|
||[parentTable](/javascript/api/word/word.table#parentTable)|获取包含此 table 的 table。|
||[parentTableCell](/javascript/api/word/word.table#parentTableCell)|获取包含此 table 的 tableCell。|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parentTableCellOrNullObject)|获取包含此 table 的 tableCell。|
||[parentTableOrNullObject](/javascript/api/word/word.table#parentTableOrNullObject)|获取包含此 table 的 table。|
||[rowCount](/javascript/api/word/word.table#rowCount)|获取表格中的行数。|
||[rows](/javascript/api/word/word.table#rows)|获取所有表格行。|
||[search (searchText： string， searchOptions？： Word.SearchOptions \| { ignorePunct？： boolean ignoreSpace？： boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： boolean matchWholeWord？： boolean matchWildcards？： boolean }) ](/javascript/api/word/word.table#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|使用指定的 SearchOptions 对 table 对象的范围执行搜索。|
||[select (selectionMode？： Word.SelectionMode) ](/javascript/api/word/word.table#select_selectionMode_)|选择表格或其开头或结尾位置，然后将 Word UI 导航到相应位置。|
||[setCellPadding (cellPaddingLocation： Word.CellPaddingLocation， cellPadding： number) ](/javascript/api/word/word.table#setCellPadding_cellPaddingLocation__cellPadding_)|设置单元格填充（以磅为单位）。|
||[shadingColor](/javascript/api/word/word.table#shadingColor)|获取并设置底纹色。|
||[style](/javascript/api/word/word.table#style)|获取或设置 table 的样式名称。|
||[styleBandedColumns](/javascript/api/word/word.table#styleBandedColumns)|获取并设置 table 是否有镶边列。|
||[styleBandedRows](/javascript/api/word/word.table#styleBandedRows)|获取并设置 table 是否有镶边行。|
||[styleBuiltIn](/javascript/api/word/word.table#styleBuiltIn)|获取或设置 table 的嵌入样式名称。|
||[styleFirstColumn](/javascript/api/word/word.table#styleFirstColumn)|获取并设置 table 的第一列是否采用特殊样式。|
||[styleLastColumn](/javascript/api/word/word.table#styleLastColumn)|获取并设置 table 的最后一列是否采用特殊样式。|
||[styleTotalRow](/javascript/api/word/word.table#styleTotalRow)|获取并设置 table 的总计行（最后一行）是否采用特殊样式。|
||[表](/javascript/api/word/word.table#tables)|获取嵌套一级的子 table。|
||[values](/javascript/api/word/word.table#values)|以 2D Javascript 数组形式获取并设置 table 中的文本值。|
||[verticalAlignment](/javascript/api/word/word.table#verticalAlignment)|获取并设置 table 中每个单元格的垂直对齐方式。|
||[width](/javascript/api/word/word.table#width)|获取并设置 table 的宽度（以磅为单位）。|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|获取或设置表边框颜色。|
||[type](/javascript/api/word/word.tableborder#type)|获取或设置表边框的类型。|
||[width](/javascript/api/word/word.tableborder#width)|获取或设置表边框的宽度（以磅为单位）。|
|[TableCell](/javascript/api/word/word.tablecell)|[body](/javascript/api/word/word.tablecell#body)|获取单元格的 body 对象。|
||[cellIndex](/javascript/api/word/word.tablecell#cellIndex)|获取单元格行中的单元格索引。|
||[columnWidth](/javascript/api/word/word.tablecell#columnWidth)|获取并设置单元格的列宽度（以磅为单位）。|
||[deleteColumn () ](/javascript/api/word/word.tablecell#deleteColumn__)|删除包含该单元格的列。|
||[deleteRow () ](/javascript/api/word/word.tablecell#deleteRow__)|删除包含该单元格的行。|
||[getBorder (borderLocation：Word.BorderLocation) ](/javascript/api/word/word.tablecell#getBorder_borderLocation_)|获取指定边框的边框样式。|
||[getCellPadding (cellPaddingLocation：Word.CellPaddingLocation) ](/javascript/api/word/word.tablecell#getCellPadding_cellPaddingLocation_)|获取单元格填充（以磅为单位）。|
||[getNext () ](/javascript/api/word/word.tablecell#getNext__)|获取下一个单元格。|
||[getNextOrNullObject () ](/javascript/api/word/word.tablecell#getNextOrNullObject__)|获取下一个单元格。|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalAlignment)|获取并设置单元格的水平对齐方式。|
||[insertColumns (insertLocation： Word.InsertLocation， columnCount： number， values？： string[][]) ](/javascript/api/word/word.tablecell#insertColumns_insertLocation__columnCount__values_)|使用单元格的列作为模板，在单元格的左侧或右侧添加列。|
||[insertRows (insertLocation： Word.InsertLocation， rowCount： number， values？： string[][]) ](/javascript/api/word/word.tablecell#insertRows_insertLocation__rowCount__values_)|使用单元格的行作为模板，在单元格的上方或下方插入行。|
||[parentRow](/javascript/api/word/word.tablecell#parentRow)|获取单元格的父行。|
||[parentTable](/javascript/api/word/word.tablecell#parentTable)|获取单元格的父表。|
||[rowIndex](/javascript/api/word/word.tablecell#rowIndex)|获取表中单元格行的索引。|
||[setCellPadding (cellPaddingLocation： Word.CellPaddingLocation， cellPadding： number) ](/javascript/api/word/word.tablecell#setCellPadding_cellPaddingLocation__cellPadding_)|设置单元格填充（以磅为单位）。|
||[shadingColor](/javascript/api/word/word.tablecell#shadingColor)|获取或设置单元格的底纹色。|
||[value](/javascript/api/word/word.tablecell#value)|获取并设置单元格的文本。|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalAlignment)|获取并设置单元格的垂直对齐方式。|
||[width](/javascript/api/word/word.tablecell#width)|获取单元格的宽度（以磅为单位）。|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getFirst__)|获取此集合中的第一个表格单元格。|
||[getFirstOrNullObject () ](/javascript/api/word/word.tablecellcollection#getFirstOrNullObject__)|获取此集合中的第一个表格单元格。|
||[items](/javascript/api/word/word.tablecellcollection#items)|获取此集合中已加载的子项。|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getFirst__)|获取此集合中的第一个表格。|
||[getFirstOrNullObject () ](/javascript/api/word/word.tablecollection#getFirstOrNullObject__)|获取此集合中的第一个表格。|
||[items](/javascript/api/word/word.tablecollection#items)|获取此集合中已加载的子项。|
|[TableRow](/javascript/api/word/word.tablerow)|[cellCount](/javascript/api/word/word.tablerow#cellCount)|获取行中的单元格数。|
||[cells](/javascript/api/word/word.tablerow#cells)|获取单元格。|
||[clear()](/javascript/api/word/word.tablerow#clear__)|清除行内容。|
||[delete()](/javascript/api/word/word.tablerow#delete__)|删除整行。|
||[font](/javascript/api/word/word.tablerow#font)|获取字体。|
||[getBorder (borderLocation：Word.BorderLocation) ](/javascript/api/word/word.tablerow#getBorder_borderLocation_)|获取行中单元格的边框样式。|
||[getCellPadding (cellPaddingLocation：Word.CellPaddingLocation) ](/javascript/api/word/word.tablerow#getCellPadding_cellPaddingLocation_)|获取单元格填充（以磅为单位）。|
||[getNext () ](/javascript/api/word/word.tablerow#getNext__)|获取下一行。|
||[getNextOrNullObject () ](/javascript/api/word/word.tablerow#getNextOrNullObject__)|获取下一行。|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalAlignment)|获取并设置行中每个单元格的水平对齐方式。|
||[ignorePunct](/javascript/api/word/word.tablerow#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.tablerow#ignoreSpace)||
||[insertRows (insertLocation： Word.InsertLocation， rowCount： number， values？： string[][]) ](/javascript/api/word/word.tablerow#insertRows_insertLocation__rowCount__values_)|使用此行作为模板插入行。|
||[isHeader](/javascript/api/word/word.tablerow#isHeader)|检查该行是否为标题行。|
||[matchCase](/javascript/api/word/word.tablerow#matchCase)||
||[matchPrefix](/javascript/api/word/word.tablerow#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.tablerow#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.tablerow#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.tablerow#matchWildcards)||
||[parentTable](/javascript/api/word/word.tablerow#parentTable)|获取父表。|
||[preferredHeight](/javascript/api/word/word.tablerow#preferredHeight)|获取并设置行的首选高度（以磅为单位）。|
||[rowIndex](/javascript/api/word/word.tablerow#rowIndex)|获取其父表中的行索引。|
||[search (searchText： string， searchOptions？： Word.SearchOptions \| { ignorePunct？： boolean ignoreSpace？： boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： boolean matchWholeWord？： boolean matchWildcards？： boolean }) ](/javascript/api/word/word.tablerow#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|在行范围内使用指定的 SearchOptions 执行搜索。|
||[select (selectionMode？： Word.SelectionMode) ](/javascript/api/word/word.tablerow#select_selectionMode_)|选择行，然后将 Word UI 导航到相应位置。|
||[setCellPadding (cellPaddingLocation： Word.CellPaddingLocation， cellPadding： number) ](/javascript/api/word/word.tablerow#setCellPadding_cellPaddingLocation__cellPadding_)|设置单元格填充（以磅为单位）。|
||[shadingColor](/javascript/api/word/word.tablerow#shadingColor)|获取并设置底纹色。|
||[values](/javascript/api/word/word.tablerow#values)|获取并设置行中的文本值，作为 2D Javascript 数组。|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalAlignment)|获取并设置行中单元格的垂直对齐方式。|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getFirst__)|获取此集合中的第一行。|
||[getFirstOrNullObject () ](/javascript/api/word/word.tablerowcollection#getFirstOrNullObject__)|获取此集合中的第一行。|
||[items](/javascript/api/word/word.tablerowcollection#items)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Word JavaScript API 要求集](word-api-requirement-sets.md)

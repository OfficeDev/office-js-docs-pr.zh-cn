---
title: Word JavaScript API 要求集 1.3
description: 有关 WordApi 1.3 要求集的详细信息。
ms.date: 03/09/2021
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: d9e0d450b601845d4e11e0fd74652c4e167f802c
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746031"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Word JavaScript API 1.3 的最近更新

WordApi 1.3 增加了对内容控件和文档级别设置的更多支持。

## <a name="api-list"></a>API 列表

下表列出了 Word JavaScript API 要求集 1.3 中的 API。 若要查看受 Word JavaScript API 要求集 1.3 或更早版本支持的所有 API 的 API 参考文档，请参阅要求集 [1.3](/javascript/api/word?view=word-js-1.3&preserve-view=true) 或更早版本中的 Word API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument (base64File？： string) ](/javascript/api/word/word.application#word-word-application-createdocument-member(1))|通过使用可选的 base64 编码文档文件创建新.docx文档。|
|[正文](/javascript/api/word/word.body)|[getRange (rangeLocation？：Word.RangeLocation) ](/javascript/api/word/word.body#word-word-body-getrange-member(1))|获取整个正文或正文的起点/终点，作为一个范围。|
||[insertTable (rowCount： number， columnCount： number， insertLocation： Word.InsertLocation， values？： string[][]) ](/javascript/api/word/word.body#word-word-body-inserttable-member(1))|插入包含指定行数和列数的 table。|
||[lists](/javascript/api/word/word.body#word-word-body-lists-member)|获取 body 中的一组 list 对象。|
||[parentBody](/javascript/api/word/word.body#word-word-body-parentbody-member)|获取 body 的父正文。|
||[parentBodyOrNullObject](/javascript/api/word/word.body#word-word-body-parentbodyornullobject-member)|获取 body 的父正文。|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#word-word-body-parentcontentcontrolornullobject-member)|获取包含正文的内容控件。|
||[parentSection](/javascript/api/word/word.body#word-word-body-parentsection-member)|获取 body 的父节。|
||[parentSectionOrNullObject](/javascript/api/word/word.body#word-word-body-parentsectionornullobject-member)|获取 body 的父节。|
||[styleBuiltIn](/javascript/api/word/word.body#word-word-body-stylebuiltin-member)|获取或设置 body 的嵌入样式名称。|
||[表](/javascript/api/word/word.body#word-word-body-tables-member)|获取 body 中的一组 table 对象。|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|获取 body 的类型。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange (rangeLocation？：Word.RangeLocation) ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getrange-member(1))|获取整个内容控件或内容控件的起点/终点，作为一个范围。|
||[getTextRanges (结束标记： string[]， trimSpacing？： boolean) ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-gettextranges-member(1))|使用标点符号和/或其他结束标记获取内容控件中的文本范围。|
||[insertTable (rowCount： number， columnCount： number， insertLocation： Word.InsertLocation， values？： string[][]) ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserttable-member(1))|将包含指定行数和列数的 table 插入 contentControl 中或在其旁边插入。|
||[lists](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-lists-member)|获取 contentControl 中的一组 list 对象。|
||[parentBody](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentbody-member)|获取 contentControl 的父正文。|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentcontentcontrolornullobject-member)|获取包含此内容控件的内容控件。|
||[parentTable](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttable-member)|获取包含 contentControl 的 table。|
||[parentTableCell](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttablecell-member)|获取包含 contentControl 的 tableCell。|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttablecellornullobject-member)|获取包含 contentControl 的 tableCell。|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttableornullobject-member)|获取包含 contentControl 的 table。|
||[split (delimiters： string[]， multiParagraphs？： boolean， trimDelimiters？： boolean， trimSpacing？： boolean) ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-split-member(1))|使用分隔符将内容控件拆分为各个子范围。|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-stylebuiltin-member)|获取或设置 contentControl 的嵌入样式名称。|
||[subtype](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-subtype-member)|获取 contentControl 的子类型。|
||[表](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-tables-member)|获取 contentControl 中的一组 table 对象。|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject (id： number) ](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbyidornullobject-member(1))|按其标识符获取内容控件。|
||[getByTypes (类型：Word.ContentControlType[]) ](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytypes-member(1))|获取具有指定类型和/或子类型的内容控件。|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getfirst-member(1))|获取此集合中的第一个内容控件。|
||[getFirstOrNullObject () ](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getfirstornullobject-member(1))|获取此集合中的第一个内容控件。|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#word-word-customproperty-delete-member(1))|删除 custom property 对象。|
||[key](/javascript/api/word/word.customproperty#word-word-customproperty-key-member)|获取 customProperty 的键。|
||[type](/javascript/api/word/word.customproperty#word-word-customproperty-type-member)|获取自定义属性的值类型。|
||[value](/javascript/api/word/word.customproperty#word-word-customproperty-value-member)|获取或设置自定义属性的值。|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add (key： string， value： any) ](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-add-member(1))|新建自定义属性或设置现有自定义属性。|
||[deleteAll () ](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-deleteall-member(1))|删除此集合中的所有自定义属性。|
||[getCount()](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getcount-member(1))|获取自定义属性的计数。|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getitem-member(1))|按键获取自定义属性对象（不区分大小写）。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getitemornullobject-member(1))|按键获取自定义属性对象（不区分大小写）。|
||[items](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-items-member)|获取此集合中已加载的子项。|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#word-word-document-properties-member)|获取文档的属性。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[body](/javascript/api/word/word.documentcreated#word-word-documentcreated-body-member)|获取文档的 body 对象。|
||[contentControls](/javascript/api/word/word.documentcreated#word-word-documentcreated-contentcontrols-member)|获取文档中的内容控件对象的集合。|
||[open () ](/javascript/api/word/word.documentcreated#word-word-documentcreated-open-member(1))|打开文档。|
||[properties](/javascript/api/word/word.documentcreated#word-word-documentcreated-properties-member)|获取文档的属性。|
||[save()](/javascript/api/word/word.documentcreated#word-word-documentcreated-save-member(1))|保存文档。|
||[saved](/javascript/api/word/word.documentcreated#word-word-documentcreated-saved-member)|指示是否已保存在文档中所做的更改。|
||[sections](/javascript/api/word/word.documentcreated#word-word-documentcreated-sections-member)|获取文档中 section 对象的集合。|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[applicationName](/javascript/api/word/word.documentproperties#word-word-documentproperties-applicationname-member)|获取 document 的应用程序名称。|
||[author](/javascript/api/word/word.documentproperties#word-word-documentproperties-author-member)|获取或设置 document 的作者。|
||[category](/javascript/api/word/word.documentproperties#word-word-documentproperties-category-member)|获取或设置 document 的类别。|
||[comments](/javascript/api/word/word.documentproperties#word-word-documentproperties-comments-member)|获取或设置 document 的注释。|
||[company](/javascript/api/word/word.documentproperties#word-word-documentproperties-company-member)|获取或设置 document 的公司。|
||[creationDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-creationdate-member)|获取文档的创建日期。|
||[customProperties](/javascript/api/word/word.documentproperties#word-word-documentproperties-customproperties-member)|获取 document 的一组 customProperty。|
||[format](/javascript/api/word/word.documentproperties#word-word-documentproperties-format-member)|获取或设置 document 的格式。|
||[keywords](/javascript/api/word/word.documentproperties#word-word-documentproperties-keywords-member)|获取或设置 document 的关键字。|
||[lastAuthor](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastauthor-member)|获取文档的最后一个作者。|
||[lastPrintDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastprintdate-member)|获取文档的上次打印日期。|
||[lastSaveTime](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastsavetime-member)|获取 document 的上次保存日期。|
||[manager](/javascript/api/word/word.documentproperties#word-word-documentproperties-manager-member)|获取或设置 document 的管理者。|
||[revisionNumber](/javascript/api/word/word.documentproperties#word-word-documentproperties-revisionnumber-member)|获取文档的修订号。|
||[security](/javascript/api/word/word.documentproperties#word-word-documentproperties-security-member)|获取文档的安全设置。|
||[subject](/javascript/api/word/word.documentproperties#word-word-documentproperties-subject-member)|获取或设置 document 的主题。|
||[template](/javascript/api/word/word.documentproperties#word-word-documentproperties-template-member)|获取文档的模板。|
||[title](/javascript/api/word/word.documentproperties#word-word-documentproperties-title-member)|获取或设置 document 的标题。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext () ](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getnext-member(1))|获取下一个嵌入式图像。|
||[getNextOrNullObject () ](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getnextornullobject-member(1))|获取下一个嵌入式图像。|
||[getRange (rangeLocation？：Word.RangeLocation) ](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getrange-member(1))|获取图片或图片的起点/终点，作为一个范围。|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentcontentcontrolornullobject-member)|获取包含嵌入式图像的内容控件。|
||[parentTable](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttable-member)|获取包含嵌入式图像的 table。|
||[parentTableCell](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttablecell-member)|获取包含嵌入式图像的 tableCell。|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttablecellornullobject-member)|获取包含嵌入式图像的 tableCell。|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttableornullobject-member)|获取包含嵌入式图像的 table。|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getfirst-member(1))|获取此集合中的第一个嵌入式图像。|
||[getFirstOrNullObject () ](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getfirstornullobject-member(1))|获取此集合中的第一个嵌入式图像。|
|[列表](/javascript/api/word/word.list)|[getLevelParagraphs (级别：number) ](/javascript/api/word/word.list#word-word-list-getlevelparagraphs-member(1))|获取列表中指定级别的段落。|
||[getLevelString (级别：number) ](/javascript/api/word/word.list#word-word-list-getlevelstring-member(1))|以字符串形式获取指定级别的项目符号、编号或图片。|
||[id](/javascript/api/word/word.list#word-word-list-id-member)|获取列表的 ID。|
||[insertParagraph (paragraphText： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.list#word-word-list-insertparagraph-member(1))|在指定位置插入段落。|
||[levelExistences](/javascript/api/word/word.list#word-word-list-levelexistences-member)|检查 list 中是否包含所有 9 个级别。|
||[levelTypes](/javascript/api/word/word.list#word-word-list-leveltypes-member)|获取 list 中的所有 9 个级别类型。|
||[paragraphs](/javascript/api/word/word.list#word-word-list-paragraphs-member)|获取 list 中的段落。|
||[setLevelAlignment (level： number， alignment： Word.Alignment) ](/javascript/api/word/word.list#word-word-list-setlevelalignment-member(1))|设置项目符号、编号或图片在列表中指定级别的对齐方式。|
||[setLevelBullet (level： number， listBullet： Word.ListBullet， charCode？： number， fontName？： string) ](/javascript/api/word/word.list#word-word-list-setlevelbullet-member(1))|设置 list 中指定级别的项目符号格式。|
||[setLevelIndents (level： number， textIndent： number， bulletNumberPictureIndent： number) ](/javascript/api/word/word.list#word-word-list-setlevelindents-member(1))|设置列表中指定级别的两种缩进方式。|
||[setLevelNumbering (level： number， listNumbering： Word.ListNumbering， formatString？： Array<string \| number>) ](/javascript/api/word/word.list#word-word-list-setlevelnumbering-member(1))|设置列表中指定级别的编号格式。|
||[setLevelStartingNumber (level： number， startingNumber： number) ](/javascript/api/word/word.list#word-word-list-setlevelstartingnumber-member(1))|设置 list 中指定级别的起始编号。|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getbyid-member(1))|按标识符获取列表。|
||[getByIdOrNullObject (id： number) ](/javascript/api/word/word.listcollection#word-word-listcollection-getbyidornullobject-member(1))|按标识符获取列表。|
||[getFirst()](/javascript/api/word/word.listcollection#word-word-listcollection-getfirst-member(1))|获取此集合中的第一个列表。|
||[getFirstOrNullObject () ](/javascript/api/word/word.listcollection#word-word-listcollection-getfirstornullobject-member(1))|获取此集合中的第一个列表。|
||[getItem(index: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getitem-member(1))|按列表对象在集合中的索引获取列表。|
||[items](/javascript/api/word/word.listcollection#word-word-listcollection-items-member)|获取此集合中已加载的子项。|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor (parentOnly？： boolean) ](/javascript/api/word/word.listitem#word-word-listitem-getancestor-member(1))|获取列表项目的父级或最近的上级元素（如果父级不存在的话）。|
||[getAncestorOrNullObject (parentOnly？： boolean) ](/javascript/api/word/word.listitem#word-word-listitem-getancestorornullobject-member(1))|获取列表项目的父级或最近的上级元素（如果父级不存在的话）。|
||[getDescendants (directChildrenOnly？： boolean) ](/javascript/api/word/word.listitem#word-word-listitem-getdescendants-member(1))|获取相应列表项目的所有后代列表项目。|
||[level](/javascript/api/word/word.listitem#word-word-listitem-level-member)|获取或设置 list 中项的级别。|
||[listString](/javascript/api/word/word.listitem#word-word-listitem-liststring-member)|获取字符串形式的列表项项目符号、编号或图片。|
||[siblingIndex](/javascript/api/word/word.listitem#word-word-listitem-siblingindex-member)|获取 listItem 相对于同级元素的序号。|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList (listId： number， level： number) ](/javascript/api/word/word.paragraph#word-word-paragraph-attachtolist-member(1))|将 paragraph 加入指定级别的现有 list。|
||[detachFromList () ](/javascript/api/word/word.paragraph#word-word-paragraph-detachfromlist-member(1))|如果此段落是列表项目的话，从列表中移出此段落。|
||[getNext () ](/javascript/api/word/word.paragraph#word-word-paragraph-getnext-member(1))|获取下一个段落。|
||[getNextOrNullObject () ](/javascript/api/word/word.paragraph#word-word-paragraph-getnextornullobject-member(1))|获取下一个段落。|
||[getPrevious () ](/javascript/api/word/word.paragraph#word-word-paragraph-getprevious-member(1))|获取上一个段落。|
||[getPreviousOrNullObject () ](/javascript/api/word/word.paragraph#word-word-paragraph-getpreviousornullobject-member(1))|获取上一个段落。|
||[getRange (rangeLocation？：Word.RangeLocation) ](/javascript/api/word/word.paragraph#word-word-paragraph-getrange-member(1))|获取整个段落或段落的起点/终点，作为一个范围。|
||[getTextRanges (结束标记： string[]， trimSpacing？： boolean) ](/javascript/api/word/word.paragraph#word-word-paragraph-gettextranges-member(1))|使用标点符号和/或其他结束标记获取段落中的文本范围。|
||[insertTable (rowCount： number， columnCount： number， insertLocation： Word.InsertLocation， values？： string[][]) ](/javascript/api/word/word.paragraph#word-word-paragraph-inserttable-member(1))|插入包含指定行数和列数的 table。|
||[isLastParagraph](/javascript/api/word/word.paragraph#word-word-paragraph-islastparagraph-member)|指明 paragraph 是其父正文内的最后一个段落。|
||[isListItem](/javascript/api/word/word.paragraph#word-word-paragraph-islistitem-member)|检查 paragraph 是否为 listItem。|
||[列表](/javascript/api/word/word.paragraph#word-word-paragraph-list-member)|获取 paragraph 所属的 List。|
||[listItem](/javascript/api/word/word.paragraph#word-word-paragraph-listitem-member)|获取 paragraph 的 ListItem。|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listitemornullobject-member)|获取 paragraph 的 ListItem。|
||[listOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listornullobject-member)|获取 paragraph 所属的 List。|
||[parentBody](/javascript/api/word/word.paragraph#word-word-paragraph-parentbody-member)|获取 paragraph 的父正文。|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parentcontentcontrolornullobject-member)|获取包含段落的内容控件。|
||[parentTable](/javascript/api/word/word.paragraph#word-word-paragraph-parenttable-member)|获取包含 paragraph 的 table。|
||[parentTableCell](/javascript/api/word/word.paragraph#word-word-paragraph-parenttablecell-member)|获取包含 paragraph 的 tableCell。|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parenttablecellornullobject-member)|获取包含 paragraph 的 tableCell。|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parenttableornullobject-member)|获取包含 paragraph 的 table。|
||[split (delimiters： string[]， trimDelimiters？： boolean， trimSpacing？： boolean) ](/javascript/api/word/word.paragraph#word-word-paragraph-split-member(1))|使用分隔符将段落拆分为多个子范围。|
||[startNewList () ](/javascript/api/word/word.paragraph#word-word-paragraph-startnewlist-member(1))|生成包含此 paragraph 的新 list。|
||[styleBuiltIn](/javascript/api/word/word.paragraph#word-word-paragraph-stylebuiltin-member)|获取或设置 paragraph 的嵌入样式名称。|
||[tableNestingLevel](/javascript/api/word/word.paragraph#word-word-paragraph-tablenestinglevel-member)|获取 paragraph 的表级别。|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getfirst-member(1))|获取此集合中的第一个段落。|
||[getFirstOrNullObject () ](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getfirstornullobject-member(1))|获取此集合中的第一个段落。|
||[getLast () ](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getlast-member(1))|获取此集合中的最后一个段落。|
||[getLastOrNullObject () ](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getlastornullobject-member(1))|获取此集合中的最后一个段落。|
|[范围](/javascript/api/word/word.range)|[compareLocationWith (range： Word.Range) ](/javascript/api/word/word.range#word-word-range-comparelocationwith-member(1))|比较此范围与另一范围的位置。|
||[expandTo (range： Word.Range) ](/javascript/api/word/word.range#word-word-range-expandto-member(1))|返回从此 range 进行任一方向扩展的新 range，以便覆盖另一 range。|
||[expandToOrNullObject (范围：Word.Range) ](/javascript/api/word/word.range#word-word-range-expandtoornullobject-member(1))|返回从此 range 进行任一方向扩展的新 range，以便覆盖另一 range。|
||[getHyperlinkRanges () ](/javascript/api/word/word.range#word-word-range-gethyperlinkranges-member(1))|获取相应范围内的超链接子范围。|
||[getNextTextRange (endingMarks： string[]， trimSpacing？： boolean) ](/javascript/api/word/word.range#word-word-range-getnexttextrange-member(1))|使用标点符号和/或其他结束标记获取下一个文本范围。|
||[getNextTextRangeOrNullObject (endingMarks： string[]， trimSpacing？： boolean) ](/javascript/api/word/word.range#word-word-range-getnexttextrangeornullobject-member(1))|使用标点符号和/或其他结束标记获取下一个文本范围。|
||[getRange (rangeLocation？：Word.RangeLocation) ](/javascript/api/word/word.range#word-word-range-getrange-member(1))|克隆相应范围，或获取该范围的起点/终点作为一个新范围。|
||[getTextRanges (结束标记： string[]， trimSpacing？： boolean) ](/javascript/api/word/word.range#word-word-range-gettextranges-member(1))|使用标点符号和/或其他结束标记获取范围中的文本子范围。|
||[hyperlink](/javascript/api/word/word.range#word-word-range-hyperlink-member)|获取 range 内的第一个超链接，或在 range 内设置超链接。|
||[insertTable (rowCount： number， columnCount： number， insertLocation： Word.InsertLocation， values？： string[][]) ](/javascript/api/word/word.range#word-word-range-inserttable-member(1))|插入包含指定行数和列数的 table。|
||[intersectWith (range： Word.Range) ](/javascript/api/word/word.range#word-word-range-intersectwith-member(1))|返回新 range 作为此 range 与另一 range 的交集。|
||[intersectWithOrNullObject (range：Word.Range) ](/javascript/api/word/word.range#word-word-range-intersectwithornullobject-member(1))|返回新 range 作为此 range 与另一 range 的交集。|
||[isEmpty](/javascript/api/word/word.range#word-word-range-isempty-member)|检查 range 长度是否为零。|
||[lists](/javascript/api/word/word.range#word-word-range-lists-member)|获取 range 中的一组 list 对象。|
||[parentBody](/javascript/api/word/word.range#word-word-range-parentbody-member)|获取 range 的父正文。|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#word-word-range-parentcontentcontrolornullobject-member)|获取包含该范围的内容控件。|
||[parentTable](/javascript/api/word/word.range#word-word-range-parenttable-member)|获取包含 range 的 table。|
||[parentTableCell](/javascript/api/word/word.range#word-word-range-parenttablecell-member)|获取包含 range 的 tableCell。|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#word-word-range-parenttablecellornullobject-member)|获取包含 range 的 tableCell。|
||[parentTableOrNullObject](/javascript/api/word/word.range#word-word-range-parenttableornullobject-member)|获取包含 range 的 table。|
||[split (delimiters： string[]， multiParagraphs？： boolean， trimDelimiters？： boolean， trimSpacing？： boolean) ](/javascript/api/word/word.range#word-word-range-split-member(1))|使用分隔符将相应范围拆分为各个子范围。|
||[styleBuiltIn](/javascript/api/word/word.range#word-word-range-stylebuiltin-member)|获取或设置 range 的嵌入样式名称。|
||[表](/javascript/api/word/word.range#word-word-range-tables-member)|获取 range 中的一组 table 对象。|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#word-word-rangecollection-getfirst-member(1))|获取此集合中的第一个范围。|
||[getFirstOrNullObject () ](/javascript/api/word/word.rangecollection#word-word-rangecollection-getfirstornullobject-member(1))|获取此集合中的第一个范围。|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#word-word-requestcontext-application-member)|[Api 集：WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext () ](/javascript/api/word/word.section#word-word-section-getnext-member(1))|获取下一节。|
||[getNextOrNullObject () ](/javascript/api/word/word.section#word-word-section-getnextornullobject-member(1))|获取下一节。|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getfirst-member(1))|获取此集合中的第一节。|
||[getFirstOrNullObject () ](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getfirstornullobject-member(1))|获取此集合中的第一节。|
|[Table](/javascript/api/word/word.table)|[addColumns (insertLocation： Word.InsertLocation， columnCount： number， values？： string[][]) ](/javascript/api/word/word.table#word-word-table-addcolumns-member(1))|使用第一个或最后一个现有列作为模板，将列添加到 table 的开头或结尾。|
||[addRows (insertLocation： Word.InsertLocation， rowCount： number， values？： string[][]) ](/javascript/api/word/word.table#word-word-table-addrows-member(1))|使用第一个或最后一个现有行作为模板，将行添加到 table 的开头或结尾。|
||[alignment](/javascript/api/word/word.table#word-word-table-alignment-member)|获取或设置表格与页面列的对齐方式。|
||[autoFitWindow () ](/javascript/api/word/word.table#word-word-table-autofitwindow-member(1))|自动调整表列，以适应窗口的宽度。|
||[clear()](/javascript/api/word/word.table#word-word-table-clear-member(1))|清除表内容。|
||[delete()](/javascript/api/word/word.table#word-word-table-delete-member(1))|删除整个表格。|
||[deleteColumns (columnIndex： number， columnCount？： number) ](/javascript/api/word/word.table#word-word-table-deletecolumns-member(1))|删除特定的列。|
||[deleteRows (rowIndex： number， rowCount？： number) ](/javascript/api/word/word.table#word-word-table-deleterows-member(1))|删除特定的行。|
||[distributeColumns () ](/javascript/api/word/word.table#word-word-table-distributecolumns-member(1))|将列设置为等宽。|
||[font](/javascript/api/word/word.table#word-word-table-font-member)|获取字体。|
||[getBorder (borderLocation：Word.BorderLocation) ](/javascript/api/word/word.table#word-word-table-getborder-member(1))|获取指定边框的边框样式。|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#word-word-table-getcell-member(1))|获取指定行和列处的表单元格。|
||[getCellOrNullObject (rowIndex： number， cellIndex： number) ](/javascript/api/word/word.table#word-word-table-getcellornullobject-member(1))|获取指定行和列处的表单元格。|
||[getCellPadding (cellPaddingLocation：Word.CellPaddingLocation) ](/javascript/api/word/word.table#word-word-table-getcellpadding-member(1))|获取单元格填充（以磅为单位）。|
||[getNext () ](/javascript/api/word/word.table#word-word-table-getnext-member(1))|获取下一个表格。|
||[getNextOrNullObject () ](/javascript/api/word/word.table#word-word-table-getnextornullobject-member(1))|获取下一个表格。|
||[getParagraphAfter () ](/javascript/api/word/word.table#word-word-table-getparagraphafter-member(1))|获取 table 之后的 paragraph。|
||[getParagraphAfterOrNullObject () ](/javascript/api/word/word.table#word-word-table-getparagraphafterornullobject-member(1))|获取 table 之后的 paragraph。|
||[getParagraphBefore () ](/javascript/api/word/word.table#word-word-table-getparagraphbefore-member(1))|获取 table 之前的 paragraph。|
||[getParagraphBeforeOrNullObject () ](/javascript/api/word/word.table#word-word-table-getparagraphbeforeornullobject-member(1))|获取 table 之前的 paragraph。|
||[getRange (rangeLocation？：Word.RangeLocation) ](/javascript/api/word/word.table#word-word-table-getrange-member(1))|获取包含此表格的范围，或包含此表格的开头或结尾的范围。|
||[headerRowCount](/javascript/api/word/word.table#word-word-table-headerrowcount-member)|获取并设置标题行数。|
||[horizontalAlignment](/javascript/api/word/word.table#word-word-table-horizontalalignment-member)|获取并设置 table 中每个单元格的水平对齐方式。|
||[ignorePunct](/javascript/api/word/word.table#word-word-table-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.table#word-word-table-ignorespace-member)||
||[insertContentControl()](/javascript/api/word/word.table#word-word-table-insertcontentcontrol-member(1))|在表格中插入内容控件。|
||[insertParagraph (paragraphText： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.table#word-word-table-insertparagraph-member(1))|在指定位置插入段落。|
||[insertTable (rowCount： number， columnCount： number， insertLocation： Word.InsertLocation， values？： string[][]) ](/javascript/api/word/word.table#word-word-table-inserttable-member(1))|插入包含指定行数和列数的 table。|
||[isUniform](/javascript/api/word/word.table#word-word-table-isuniform-member)|指明所有表行是否一致。|
||[matchCase](/javascript/api/word/word.table#word-word-table-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.table#word-word-table-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.table#word-word-table-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.table#word-word-table-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.table#word-word-table-matchwildcards-member)||
||[nestingLevel](/javascript/api/word/word.table#word-word-table-nestinglevel-member)|获取 table 的嵌套级别。|
||[parentBody](/javascript/api/word/word.table#word-word-table-parentbody-member)|获取 table 的父正文。|
||[parentContentControl](/javascript/api/word/word.table#word-word-table-parentcontentcontrol-member)|获取包含 table 的 contentControl。|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#word-word-table-parentcontentcontrolornullobject-member)|获取包含 table 的 contentControl。|
||[parentTable](/javascript/api/word/word.table#word-word-table-parenttable-member)|获取包含此 table 的 table。|
||[parentTableCell](/javascript/api/word/word.table#word-word-table-parenttablecell-member)|获取包含此 table 的 tableCell。|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#word-word-table-parenttablecellornullobject-member)|获取包含此 table 的 tableCell。|
||[parentTableOrNullObject](/javascript/api/word/word.table#word-word-table-parenttableornullobject-member)|获取包含此 table 的 table。|
||[rowCount](/javascript/api/word/word.table#word-word-table-rowcount-member)|获取表格中的行数。|
||[rows](/javascript/api/word/word.table#word-word-table-rows-member)|获取所有表格行。|
||[search (searchText： string， searchOptions？： Word.SearchOptions \| { ignorePunct？： boolean ignoreSpace？： boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： boolean matchWholeWord？： boolean matchWildcards？： boolean }) ](/javascript/api/word/word.table#word-word-table-search-member(1))|使用指定的 SearchOptions 对 table 对象的范围执行搜索。|
||[select (selectionMode？： Word.SelectionMode) ](/javascript/api/word/word.table#word-word-table-select-member(1))|选择表格或其开头或结尾位置，然后将 Word UI 导航到相应位置。|
||[setCellPadding (cellPaddingLocation： Word.CellPaddingLocation， cellPadding： number) ](/javascript/api/word/word.table#word-word-table-setcellpadding-member(1))|设置单元格填充（以磅为单位）。|
||[shadingColor](/javascript/api/word/word.table#word-word-table-shadingcolor-member)|获取并设置底纹色。|
||[style](/javascript/api/word/word.table#word-word-table-style-member)|获取或设置 table 的样式名称。|
||[styleBandedColumns](/javascript/api/word/word.table#word-word-table-stylebandedcolumns-member)|获取并设置 table 是否有镶边列。|
||[styleBandedRows](/javascript/api/word/word.table#word-word-table-stylebandedrows-member)|获取并设置 table 是否有镶边行。|
||[styleBuiltIn](/javascript/api/word/word.table#word-word-table-stylebuiltin-member)|获取或设置 table 的嵌入样式名称。|
||[styleFirstColumn](/javascript/api/word/word.table#word-word-table-stylefirstcolumn-member)|获取并设置 table 的第一列是否采用特殊样式。|
||[styleLastColumn](/javascript/api/word/word.table#word-word-table-stylelastcolumn-member)|获取并设置 table 的最后一列是否采用特殊样式。|
||[styleTotalRow](/javascript/api/word/word.table#word-word-table-styletotalrow-member)|获取并设置 table 的总计行（最后一行）是否采用特殊样式。|
||[表](/javascript/api/word/word.table#word-word-table-tables-member)|获取嵌套一级的子 table。|
||[values](/javascript/api/word/word.table#word-word-table-values-member)|以 2D Javascript 数组形式获取并设置 table 中的文本值。|
||[verticalAlignment](/javascript/api/word/word.table#word-word-table-verticalalignment-member)|获取并设置 table 中每个单元格的垂直对齐方式。|
||[width](/javascript/api/word/word.table#word-word-table-width-member)|获取并设置 table 的宽度（以磅为单位）。|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#word-word-tableborder-color-member)|获取或设置表边框颜色。|
||[type](/javascript/api/word/word.tableborder#word-word-tableborder-type-member)|获取或设置表边框的类型。|
||[width](/javascript/api/word/word.tableborder#word-word-tableborder-width-member)|获取或设置表边框的宽度（以磅为单位）。|
|[TableCell](/javascript/api/word/word.tablecell)|[body](/javascript/api/word/word.tablecell#word-word-tablecell-body-member)|获取单元格的 body 对象。|
||[cellIndex](/javascript/api/word/word.tablecell#word-word-tablecell-cellindex-member)|获取单元格行中的单元格索引。|
||[columnWidth](/javascript/api/word/word.tablecell#word-word-tablecell-columnwidth-member)|获取并设置单元格的列宽度（以磅为单位）。|
||[deleteColumn () ](/javascript/api/word/word.tablecell#word-word-tablecell-deletecolumn-member(1))|删除包含该单元格的列。|
||[deleteRow () ](/javascript/api/word/word.tablecell#word-word-tablecell-deleterow-member(1))|删除包含该单元格的行。|
||[getBorder (borderLocation：Word.BorderLocation) ](/javascript/api/word/word.tablecell#word-word-tablecell-getborder-member(1))|获取指定边框的边框样式。|
||[getCellPadding (cellPaddingLocation：Word.CellPaddingLocation) ](/javascript/api/word/word.tablecell#word-word-tablecell-getcellpadding-member(1))|获取单元格填充（以磅为单位）。|
||[getNext () ](/javascript/api/word/word.tablecell#word-word-tablecell-getnext-member(1))|获取下一个单元格。|
||[getNextOrNullObject () ](/javascript/api/word/word.tablecell#word-word-tablecell-getnextornullobject-member(1))|获取下一个单元格。|
||[horizontalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-horizontalalignment-member)|获取并设置单元格的水平对齐方式。|
||[insertColumns (insertLocation： Word.InsertLocation， columnCount： number， values？： string[][]) ](/javascript/api/word/word.tablecell#word-word-tablecell-insertcolumns-member(1))|使用单元格的列作为模板，在单元格的左侧或右侧添加列。|
||[insertRows (insertLocation： Word.InsertLocation， rowCount： number， values？： string[][]) ](/javascript/api/word/word.tablecell#word-word-tablecell-insertrows-member(1))|使用单元格的行作为模板，在单元格的上方或下方插入行。|
||[parentRow](/javascript/api/word/word.tablecell#word-word-tablecell-parentrow-member)|获取单元格的父行。|
||[parentTable](/javascript/api/word/word.tablecell#word-word-tablecell-parenttable-member)|获取单元格的父表。|
||[rowIndex](/javascript/api/word/word.tablecell#word-word-tablecell-rowindex-member)|获取表中单元格行的索引。|
||[setCellPadding (cellPaddingLocation： Word.CellPaddingLocation， cellPadding： number) ](/javascript/api/word/word.tablecell#word-word-tablecell-setcellpadding-member(1))|设置单元格填充（以磅为单位）。|
||[shadingColor](/javascript/api/word/word.tablecell#word-word-tablecell-shadingcolor-member)|获取或设置单元格的底纹色。|
||[value](/javascript/api/word/word.tablecell#word-word-tablecell-value-member)|获取并设置单元格的文本。|
||[verticalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-verticalalignment-member)|获取并设置单元格的垂直对齐方式。|
||[width](/javascript/api/word/word.tablecell#word-word-tablecell-width-member)|获取单元格的宽度（以磅为单位）。|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getfirst-member(1))|获取此集合中的第一个表格单元格。|
||[getFirstOrNullObject () ](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getfirstornullobject-member(1))|获取此集合中的第一个表格单元格。|
||[items](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-items-member)|获取此集合中已加载的子项。|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#word-word-tablecollection-getfirst-member(1))|获取此集合中的第一个表格。|
||[getFirstOrNullObject () ](/javascript/api/word/word.tablecollection#word-word-tablecollection-getfirstornullobject-member(1))|获取此集合中的第一个表格。|
||[items](/javascript/api/word/word.tablecollection#word-word-tablecollection-items-member)|获取此集合中已加载的子项。|
|[TableRow](/javascript/api/word/word.tablerow)|[cellCount](/javascript/api/word/word.tablerow#word-word-tablerow-cellcount-member)|获取行中的单元格数。|
||[cells](/javascript/api/word/word.tablerow#word-word-tablerow-cells-member)|获取单元格。|
||[clear()](/javascript/api/word/word.tablerow#word-word-tablerow-clear-member(1))|清除行内容。|
||[delete()](/javascript/api/word/word.tablerow#word-word-tablerow-delete-member(1))|删除整行。|
||[font](/javascript/api/word/word.tablerow#word-word-tablerow-font-member)|获取字体。|
||[getBorder (borderLocation：Word.BorderLocation) ](/javascript/api/word/word.tablerow#word-word-tablerow-getborder-member(1))|获取行中单元格的边框样式。|
||[getCellPadding (cellPaddingLocation：Word.CellPaddingLocation) ](/javascript/api/word/word.tablerow#word-word-tablerow-getcellpadding-member(1))|获取单元格填充（以磅为单位）。|
||[getNext () ](/javascript/api/word/word.tablerow#word-word-tablerow-getnext-member(1))|获取下一行。|
||[getNextOrNullObject () ](/javascript/api/word/word.tablerow#word-word-tablerow-getnextornullobject-member(1))|获取下一行。|
||[horizontalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-horizontalalignment-member)|获取并设置行中每个单元格的水平对齐方式。|
||[ignorePunct](/javascript/api/word/word.tablerow#word-word-tablerow-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.tablerow#word-word-tablerow-ignorespace-member)||
||[insertRows (insertLocation： Word.InsertLocation， rowCount： number， values？： string[][]) ](/javascript/api/word/word.tablerow#word-word-tablerow-insertrows-member(1))|使用此行作为模板插入行。|
||[isHeader](/javascript/api/word/word.tablerow#word-word-tablerow-isheader-member)|检查该行是否为标题行。|
||[matchCase](/javascript/api/word/word.tablerow#word-word-tablerow-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.tablerow#word-word-tablerow-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.tablerow#word-word-tablerow-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.tablerow#word-word-tablerow-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.tablerow#word-word-tablerow-matchwildcards-member)||
||[parentTable](/javascript/api/word/word.tablerow#word-word-tablerow-parenttable-member)|获取父表。|
||[preferredHeight](/javascript/api/word/word.tablerow#word-word-tablerow-preferredheight-member)|获取并设置行的首选高度（以磅为单位）。|
||[rowIndex](/javascript/api/word/word.tablerow#word-word-tablerow-rowindex-member)|获取其父表中的行索引。|
||[search (searchText： string， searchOptions？： Word.SearchOptions \| { ignorePunct？： boolean ignoreSpace？： boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： boolean matchWholeWord？： boolean matchWildcards？： boolean }) ](/javascript/api/word/word.tablerow#word-word-tablerow-search-member(1))|在行范围内使用指定的 SearchOptions 执行搜索。|
||[select (selectionMode？： Word.SelectionMode) ](/javascript/api/word/word.tablerow#word-word-tablerow-select-member(1))|选择行，然后将 Word UI 导航到相应位置。|
||[setCellPadding (cellPaddingLocation： Word.CellPaddingLocation， cellPadding： number) ](/javascript/api/word/word.tablerow#word-word-tablerow-setcellpadding-member(1))|设置单元格填充（以磅为单位）。|
||[shadingColor](/javascript/api/word/word.tablerow#word-word-tablerow-shadingcolor-member)|获取并设置底纹色。|
||[values](/javascript/api/word/word.tablerow#word-word-tablerow-values-member)|获取并设置行中的文本值，作为 2D Javascript 数组。|
||[verticalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-verticalalignment-member)|获取并设置行中单元格的垂直对齐方式。|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getfirst-member(1))|获取此集合中的第一行。|
||[getFirstOrNullObject () ](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getfirstornullobject-member(1))|获取此集合中的第一行。|
||[items](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-items-member)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Word JavaScript API 要求集](word-api-requirement-sets.md)

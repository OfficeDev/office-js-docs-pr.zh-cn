---
title: Word JavaScript API 要求集1。3
description: 有关 WordApi 1.3 要求集的详细信息
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 1344d66f2a4d9a3c9ff93c042fa1f23013e1bb27
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996428"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Word JavaScript API 1.3 的最近更新

WordApi 1.3 增加了对内容控件、自定义 XML 和文档级设置的更多支持。

## <a name="api-list"></a>API 列表

下表列出了 Word JavaScript API 要求集1.3 中的 Api。 若要查看 Word JavaScript API 要求集1.3 或更早版本支持的所有 Api 的 API 参考文档，请参阅 [要求集1.3 或更早版本中的 Word api](/javascript/api/word?view=word-js-1.3&preserve-view=true)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument (base64File？： string) ](/javascript/api/word/word.application#createdocument-base64file-)|使用可选的 base64 编码的 .docx 文件创建一个新文档。|
|[正文](/javascript/api/word/word.body)|[getRange (rangeLocation？： Word RangeLocation) ](/javascript/api/word/word.body#getrange-rangelocation-)|获取整个正文或正文的起点/终点，作为一个范围。|
||[insertTable (rowCount： number，columnCount： number，insertLocation： InsertLocation，values？： string [] [] ) ](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|插入包含指定行数和列数的 table。|
||[lists](/javascript/api/word/word.body#lists)|获取 body 中的一组 list 对象。|
||[parentBody](/javascript/api/word/word.body#parentbody)|获取 body 的父正文。|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|获取 body 的父正文。|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|获取包含正文的内容控件。|
||[parentSection](/javascript/api/word/word.body#parentsection)|获取 body 的父节。|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|获取 body 的父节。|
||[表](/javascript/api/word/word.body#tables)|获取 body 中的一组 table 对象。|
||[type](/javascript/api/word/word.body#type)|获取 body 的类型。|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|获取或设置 body 的嵌入样式名称。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange (rangeLocation？： Word RangeLocation) ](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|获取整个内容控件或内容控件的起点/终点，作为一个范围。|
||[getTextRanges (endingMarks： string []，trimSpacing？： boolean) ](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|使用标点符号和/或其他结束标记获取内容控件中的文本范围。|
||[insertTable (rowCount： number，columnCount： number，insertLocation： InsertLocation，values？： string [] [] ) ](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|将包含指定行数和列数的 table 插入 contentControl 中或在其旁边插入。|
||[lists](/javascript/api/word/word.contentcontrol#lists)|获取 contentControl 中的一组 list 对象。|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|获取 contentControl 的父正文。|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|获取包含此内容控件的内容控件。|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|获取包含 contentControl 的 table。|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|获取包含 contentControl 的 tableCell。|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|获取包含 contentControl 的 tableCell。|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parenttableornullobject)|获取包含 contentControl 的 table。|
||[个子](/javascript/api/word/word.contentcontrol#subtype)|获取 contentControl 的子类型。|
||[表](/javascript/api/word/word.contentcontrol#tables)|获取 contentControl 中的一组 table 对象。|
||[拆分 (定界符： string []，multiParagraphs？： boolean，trimDelimiters？： boolean，trimSpacing？： boolean) ](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|使用分隔符将内容控件拆分为各个子范围。|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|获取或设置 contentControl 的嵌入样式名称。|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject (id： number) ](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|按其标识符获取内容控件。|
||[getByTypes (类型： ContentControlType [] ) ](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|获取具有指定类型和/或子类型的内容控件。|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|获取此集合中的第一个内容控件。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|获取此集合中的第一个内容控件。|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|删除 custom property 对象。|
||[key](/javascript/api/word/word.customproperty#key)|获取 customProperty 的键。|
||[type](/javascript/api/word/word.customproperty#type)|获取自定义属性的值类型。|
||[value](/javascript/api/word/word.customproperty#value)|获取或设置自定义属性的值。|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add (key： string，value： any) ](/javascript/api/word/word.custompropertycollection#add-key--value-)|新建自定义属性或设置现有自定义属性。|
||[deleteAll ( # B1 ](/javascript/api/word/word.custompropertycollection#deleteall--)|删除此集合中的所有自定义属性。|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|获取自定义属性的计数。|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|按键获取自定义属性对象（不区分大小写）。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|按键获取自定义属性对象（不区分大小写）。|
||[items](/javascript/api/word/word.custompropertycollection#items)|获取此集合中已加载的子项。|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|获取文档的属性。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[打开 ( # B1 ](/javascript/api/word/word.documentcreated#open--)|打开文档。|
||[body](/javascript/api/word/word.documentcreated#body)|获取文档的正文对象。|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|获取文档中的内容控件对象的集合。|
||[properties](/javascript/api/word/word.documentcreated#properties)|获取文档的属性。|
||[保存](/javascript/api/word/word.documentcreated#saved)|指示是否已保存在文档中所做的更改。|
||[sections](/javascript/api/word/word.documentcreated#sections)|获取文档中的节对象的集合。|
||[save()](/javascript/api/word/word.documentcreated#save--)|保存文档。|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[编写](/javascript/api/word/word.documentproperties#author)|获取或设置 document 的作者。|
||[类别](/javascript/api/word/word.documentproperties#category)|获取或设置 document 的类别。|
||[comments](/javascript/api/word/word.documentproperties#comments)|获取或设置 document 的注释。|
||[company](/javascript/api/word/word.documentproperties#company)|获取或设置 document 的公司。|
||[format](/javascript/api/word/word.documentproperties#format)|获取或设置 document 的格式。|
||[关键字](/javascript/api/word/word.documentproperties#keywords)|获取或设置 document 的关键字。|
||[manager](/javascript/api/word/word.documentproperties#manager)|获取或设置 document 的管理者。|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|获取 document 的应用程序名称。|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|获取文档的创建日期。|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|获取 document 的一组 customProperty。|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|获取文档的最后一个作者。|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastprintdate)|获取文档的上次打印日期。|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|获取 document 的上次保存日期。|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|获取文档的修订号。|
||[保护](/javascript/api/word/word.documentproperties#security)|获取文档的安全设置。|
||[template](/javascript/api/word/word.documentproperties#template)|获取文档的模板。|
||[subject](/javascript/api/word/word.documentproperties#subject)|获取或设置 document 的主题。|
||[title](/javascript/api/word/word.documentproperties#title)|获取或设置 document 的标题。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext ( # B1 ](/javascript/api/word/word.inlinepicture#getnext--)|获取下一个嵌入式图像。|
||[getNextOrNullObject ( # B1 ](/javascript/api/word/word.inlinepicture#getnextornullobject--)|获取下一个嵌入式图像。|
||[getRange (rangeLocation？： Word RangeLocation) ](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|获取图片或图片的起点/终点，作为一个范围。|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|获取包含嵌入式图像的内容控件。|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|获取包含嵌入式图像的 table。|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|获取包含嵌入式图像的 tableCell。|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|获取包含嵌入式图像的 tableCell。|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parenttableornullobject)|获取包含嵌入式图像的 table。|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|获取此集合中的第一个嵌入式图像。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|获取此集合中的第一个嵌入式图像。|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs (级别： number) ](/javascript/api/word/word.list#getlevelparagraphs-level-)|获取列表中指定级别的段落。|
||[getLevelString (级别： number) ](/javascript/api/word/word.list#getlevelstring-level-)|以字符串形式获取指定级别的项目符号、编号或图片。|
||[insertParagraph (paragraphText： string，insertLocation： InsertLocation) ](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|在指定位置插入段落。|
||[id](/javascript/api/word/word.list#id)|获取列表的 id。|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|检查 list 中是否包含所有 9 个级别。|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|获取 list 中的所有 9 个级别类型。|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|获取 list 中的段落。|
||[setLevelAlignment (level： number，对齐方式： Word) 对齐方式 ](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|设置列表中指定级别的项目符号、编号或图片的对齐方式。|
||[setLevelBullet (level： number，listBullet： ListBullet，charCode？： number，fontName？： string) ](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|设置 list 中指定级别的项目符号格式。|
||[setLevelIndents (level： number，textIndent： number，bulletNumberPictureIndent： number) ](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|设置列表中指定级别的两种缩进方式。|
||[setLevelNumbering (level： number，listNumbering： ListNumbering，格式说明符？： Array<string \| number>) ](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|设置列表中指定级别的编号格式。|
||[setLevelStartingNumber (level： number，startingNumber： number) ](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|设置 list 中指定级别的起始编号。|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|按标识符获取列表。|
||[getByIdOrNullObject (id： number) ](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|按标识符获取列表。|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|获取此集合中的第一个列表。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.listcollection#getfirstornullobject--)|获取此集合中的第一个列表。|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getitem-index-)|按列表对象在集合中的索引获取列表。|
||[items](/javascript/api/word/word.listcollection#items)|获取此集合中已加载的子项。|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor (parentOnly？： boolean) ](/javascript/api/word/word.listitem#getancestor-parentonly-)|获取列表项目的父级或最近的上级元素（如果父级不存在的话）。|
||[getAncestorOrNullObject (parentOnly？： boolean) ](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|获取列表项目的父级或最近的上级元素（如果父级不存在的话）。|
||[getDescendants (directChildrenOnly？： boolean) ](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|获取相应列表项目的所有后代列表项目。|
||[level](/javascript/api/word/word.listitem#level)|获取或设置 list 中项的级别。|
||[listString](/javascript/api/word/word.listitem#liststring)|以字符串形式获取列表项的项目符号、编号或图片。|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|获取 listItem 相对于同级元素的序号。|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList (listId： number，level： number) ](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|将 paragraph 加入指定级别的现有 list。|
||[detachFromList ( # B1 ](/javascript/api/word/word.paragraph#detachfromlist--)|如果此段落是列表项目的话，从列表中移出此段落。|
||[getNext ( # B1 ](/javascript/api/word/word.paragraph#getnext--)|获取下一个段落。|
||[getNextOrNullObject ( # B1 ](/javascript/api/word/word.paragraph#getnextornullobject--)|获取下一个段落。|
||[getPrevious ( # B1 ](/javascript/api/word/word.paragraph#getprevious--)|获取上一个段落。|
||[getPreviousOrNullObject ( # B1 ](/javascript/api/word/word.paragraph#getpreviousornullobject--)|获取上一个段落。|
||[getRange (rangeLocation？： Word RangeLocation) ](/javascript/api/word/word.paragraph#getrange-rangelocation-)|获取整个段落或段落的起点/终点，作为一个范围。|
||[getTextRanges (endingMarks： string []，trimSpacing？： boolean) ](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|使用标点符号和/或其他结束标记获取段落中的文本范围。|
||[insertTable (rowCount： number，columnCount： number，insertLocation： InsertLocation，values？： string [] [] ) ](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|插入包含指定行数和列数的 table。|
||[isLastParagraph](/javascript/api/word/word.paragraph#islastparagraph)|指明 paragraph 是其父正文内的最后一个段落。|
||[isListItem](/javascript/api/word/word.paragraph#islistitem)|检查 paragraph 是否为 listItem。|
||[list](/javascript/api/word/word.paragraph#list)|获取 paragraph 所属的 List。|
||[listItem](/javascript/api/word/word.paragraph#listitem)|获取 paragraph 的 ListItem。|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listitemornullobject)|获取 paragraph 的 ListItem。|
||[listOrNullObject](/javascript/api/word/word.paragraph#listornullobject)|获取 paragraph 所属的 List。|
||[parentBody](/javascript/api/word/word.paragraph#parentbody)|获取 paragraph 的父正文。|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentcontentcontrolornullobject)|获取包含段落的内容控件。|
||[parentTable](/javascript/api/word/word.paragraph#parenttable)|获取包含 paragraph 的 table。|
||[parentTableCell](/javascript/api/word/word.paragraph#parenttablecell)|获取包含 paragraph 的 tableCell。|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parenttablecellornullobject)|获取包含 paragraph 的 tableCell。|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parenttableornullobject)|获取包含 paragraph 的 table。|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tablenestinglevel)|获取 paragraph 的表级别。|
||[拆分 (定界符： string []，trimDelimiters？： boolean，trimSpacing？： boolean) ](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|使用分隔符将段落拆分为多个子范围。|
||[startNewList ( # B1 ](/javascript/api/word/word.paragraph#startnewlist--)|生成包含此 paragraph 的新 list。|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|获取或设置 paragraph 的嵌入样式名称。|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|获取此集合中的第一个段落。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|获取此集合中的第一个段落。|
||[getLast ( # B1 ](/javascript/api/word/word.paragraphcollection#getlast--)|获取此集合中的最后一个段落。|
||[getLastOrNullObject ( # B1 ](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|获取此集合中的最后一个段落。|
|[Range](/javascript/api/word/word.range)|[compareLocationWith (范围： Word) 区域 ](/javascript/api/word/word.range#comparelocationwith-range-)|比较此范围与另一范围的位置。|
||[expandTo (范围： Word) 区域 ](/javascript/api/word/word.range#expandto-range-)|返回从此 range 进行任一方向扩展的新 range，以便覆盖另一 range。|
||[expandToOrNullObject (范围： Word) 区域 ](/javascript/api/word/word.range#expandtoornullobject-range-)|返回从此 range 进行任一方向扩展的新 range，以便覆盖另一 range。|
||[getHyperlinkRanges ( # B1 ](/javascript/api/word/word.range#gethyperlinkranges--)|获取相应范围内的超链接子范围。|
||[getNextTextRange (endingMarks： string []，trimSpacing？： boolean) ](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|使用标点符号和/或其他结束标记获取下一个文本范围。|
||[getNextTextRangeOrNullObject (endingMarks： string []，trimSpacing？： boolean) ](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|使用标点符号和/或其他结束标记获取下一个文本范围。|
||[getRange (rangeLocation？： Word RangeLocation) ](/javascript/api/word/word.range#getrange-rangelocation-)|克隆相应范围，或获取该范围的起点/终点作为一个新范围。|
||[getTextRanges (endingMarks： string []，trimSpacing？： boolean) ](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|使用标点符号和/或其他结束标记获取区域中的文本子范围。|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|获取 range 内的第一个超链接，或在 range 内设置超链接。|
||[insertTable (rowCount： number，columnCount： number，insertLocation： InsertLocation，values？： string [] [] ) ](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|插入包含指定行数和列数的 table。|
||[intersectWith (范围： Word) 区域 ](/javascript/api/word/word.range#intersectwith-range-)|返回新 range 作为此 range 与另一 range 的交集。|
||[intersectWithOrNullObject (范围： Word) 区域 ](/javascript/api/word/word.range#intersectwithornullobject-range-)|返回新 range 作为此 range 与另一 range 的交集。|
||[isEmpty](/javascript/api/word/word.range#isempty)|检查 range 长度是否为零。|
||[lists](/javascript/api/word/word.range#lists)|获取 range 中的一组 list 对象。|
||[parentBody](/javascript/api/word/word.range#parentbody)|获取 range 的父正文。|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|获取包含该范围的内容控件。|
||[parentTable](/javascript/api/word/word.range#parenttable)|获取包含 range 的 table。|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|获取包含 range 的 tableCell。|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parenttablecellornullobject)|获取包含 range 的 tableCell。|
||[parentTableOrNullObject](/javascript/api/word/word.range#parenttableornullobject)|获取包含 range 的 table。|
||[表](/javascript/api/word/word.range#tables)|获取 range 中的一组 table 对象。|
||[拆分 (定界符： string []，multiParagraphs？： boolean，trimDelimiters？： boolean，trimSpacing？： boolean) ](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|使用分隔符将相应范围拆分为各个子范围。|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|获取或设置 range 的嵌入样式名称。|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|获取此集合中的第一个范围。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.rangecollection#getfirstornullobject--)|获取此集合中的第一个范围。|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#application)|[Api set： WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext ( # B1 ](/javascript/api/word/word.section#getnext--)|获取下一节。|
||[getNextOrNullObject ( # B1 ](/javascript/api/word/word.section#getnextornullobject--)|获取下一节。|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|获取此集合中的第一节。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|获取此集合中的第一节。|
|[Table](/javascript/api/word/word.table)|[addColumns (insertLocation： InsertLocation，columnCount： number，values？： string [] [] ) ](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|使用第一个或最后一个现有列作为模板，将列添加到 table 的开头或结尾。|
||[addRows (insertLocation： InsertLocation，rowCount： number，values？： string [] [] ) ](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|使用第一个或最后一个现有行作为模板，将行添加到 table 的开头或结尾。|
||[对齐方式](/javascript/api/word/word.table#alignment)|获取或设置表相对于页列的对齐方式。|
||[autoFitWindow ( # B1 ](/javascript/api/word/word.table#autofitwindow--)|自动调整表列，以适应窗口的宽度。|
||[clear()](/javascript/api/word/word.table#clear--)|清除表内容。|
||[delete()](/javascript/api/word/word.table#delete--)|删除整个表格。|
||[deleteColumns (columnIndex： number，columnCount？： number) ](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|删除特定的列。|
||[deleteRows (rowIndex： number，rowCount？： number) ](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|删除特定的行。|
||[distributeColumns ( # B1 ](/javascript/api/word/word.table#distributecolumns--)|将列设置为等宽。|
||[getBorder (borderLocation： Word BorderLocation) ](/javascript/api/word/word.table#getborder-borderlocation-)|获取指定边框的边框样式。|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|获取指定行和列处的表单元格。|
||[getCellOrNullObject (rowIndex： number，cellIndex： number) ](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|获取指定行和列处的表单元格。|
||[getCellPadding (cellPaddingLocation： Word CellPaddingLocation) ](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|获取单元格填充（以磅为单位）。|
||[getNext ( # B1 ](/javascript/api/word/word.table#getnext--)|获取下一个表格。|
||[getNextOrNullObject ( # B1 ](/javascript/api/word/word.table#getnextornullobject--)|获取下一个表格。|
||[getParagraphAfter ( # B1 ](/javascript/api/word/word.table#getparagraphafter--)|获取 table 之后的 paragraph。|
||[getParagraphAfterOrNullObject ( # B1 ](/javascript/api/word/word.table#getparagraphafterornullobject--)|获取 table 之后的 paragraph。|
||[getParagraphBefore ( # B1 ](/javascript/api/word/word.table#getparagraphbefore--)|获取 table 之前的 paragraph。|
||[getParagraphBeforeOrNullObject ( # B1 ](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|获取 table 之前的 paragraph。|
||[getRange (rangeLocation？： Word RangeLocation) ](/javascript/api/word/word.table#getrange-rangelocation-)|获取包含此表格的范围，或包含此表格的开头或结尾的范围。|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|获取并设置标题行数。|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|获取并设置 table 中每个单元格的水平对齐方式。|
||[ignorePunct](/javascript/api/word/word.table#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.table#ignorespace)||
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|在表格中插入内容控件。|
||[insertParagraph (paragraphText： string，insertLocation： InsertLocation) ](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|在指定位置插入段落。|
||[insertTable (rowCount： number，columnCount： number，insertLocation： InsertLocation，values？： string [] [] ) ](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|插入包含指定行数和列数的 table。|
||[matchCase](/javascript/api/word/word.table#matchcase)||
||[matchPrefix](/javascript/api/word/word.table#matchprefix)||
||[matchSuffix](/javascript/api/word/word.table#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.table#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.table#matchwildcards)||
||[font](/javascript/api/word/word.table#font)|获取字体。|
||[isUniform](/javascript/api/word/word.table#isuniform)|指明所有表行是否一致。|
||[nestingLevel](/javascript/api/word/word.table#nestinglevel)|获取 table 的嵌套级别。|
||[parentBody](/javascript/api/word/word.table#parentbody)|获取 table 的父正文。|
||[parentContentControl](/javascript/api/word/word.table#parentcontentcontrol)|获取包含 table 的 contentControl。|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentcontentcontrolornullobject)|获取包含 table 的 contentControl。|
||[parentTable](/javascript/api/word/word.table#parenttable)|获取包含此 table 的 table。|
||[parentTableCell](/javascript/api/word/word.table#parenttablecell)|获取包含此 table 的 tableCell。|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parenttablecellornullobject)|获取包含此 table 的 tableCell。|
||[parentTableOrNullObject](/javascript/api/word/word.table#parenttableornullobject)|获取包含此 table 的 table。|
||[rowCount](/javascript/api/word/word.table#rowcount)|获取表格中的行数。|
||[rows](/javascript/api/word/word.table#rows)|获取所有表格行。|
||[表](/javascript/api/word/word.table#tables)|获取嵌套一级的子 table。|
||[search (searchText： string，searchOptions？： Word. SearchOptions \| {ignorePunct？： Boolean ignoreSpace？： Boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： Boolean matchWholeWord？： Boolean matchWildcards？： boolean} ) ](/javascript/api/word/word.table#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|在表对象的作用域上使用指定的 SearchOptions 执行搜索。|
||[选择 (selectionMode？： SelectionMode) ](/javascript/api/word/word.table#select-selectionmode-)|选择表格或其开头或结尾位置，然后将 Word UI 导航到相应位置。|
||[setCellPadding (cellPaddingLocation： CellPaddingLocation，cellPadding： number) ](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)|设置单元格填充（以磅为单位）。|
||[shadingColor](/javascript/api/word/word.table#shadingcolor)|获取并设置底纹色。|
||[style](/javascript/api/word/word.table#style)|获取或设置 table 的样式名称。|
||[styleBandedColumns](/javascript/api/word/word.table#stylebandedcolumns)|获取并设置 table 是否有镶边列。|
||[styleBandedRows](/javascript/api/word/word.table#stylebandedrows)|获取并设置 table 是否有镶边行。|
||[styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)|获取或设置 table 的嵌入样式名称。|
||[styleFirstColumn](/javascript/api/word/word.table#stylefirstcolumn)|获取并设置 table 的第一列是否采用特殊样式。|
||[styleLastColumn](/javascript/api/word/word.table#stylelastcolumn)|获取并设置 table 的最后一列是否采用特殊样式。|
||[styleTotalRow](/javascript/api/word/word.table#styletotalrow)|获取并设置 table 的总计行（最后一行）是否采用特殊样式。|
||[values](/javascript/api/word/word.table#values)|以 2D Javascript 数组形式获取并设置 table 中的文本值。|
||[verticalAlignment](/javascript/api/word/word.table#verticalalignment)|获取并设置 table 中每个单元格的垂直对齐方式。|
||[width](/javascript/api/word/word.table#width)|获取并设置 table 的宽度（以磅为单位）。|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|获取或设置表格边框颜色。|
||[type](/javascript/api/word/word.tableborder#type)|获取或设置表边框的类型。|
||[width](/javascript/api/word/word.tableborder#width)|获取或设置表边框的宽度（以磅为单位）。|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnwidth)|获取并设置单元格的列宽度（以磅为单位）。|
||[deleteColumn ( # B1 ](/javascript/api/word/word.tablecell#deletecolumn--)|删除包含该单元格的列。|
||[deleteRow ( # B1 ](/javascript/api/word/word.tablecell#deleterow--)|删除包含该单元格的行。|
||[getBorder (borderLocation： Word BorderLocation) ](/javascript/api/word/word.tablecell#getborder-borderlocation-)|获取指定边框的边框样式。|
||[getCellPadding (cellPaddingLocation： Word CellPaddingLocation) ](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)|获取单元格填充（以磅为单位）。|
||[getNext ( # B1 ](/javascript/api/word/word.tablecell#getnext--)|获取下一个单元格。|
||[getNextOrNullObject ( # B1 ](/javascript/api/word/word.tablecell#getnextornullobject--)|获取下一个单元格。|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)|获取并设置单元格的水平对齐方式。|
||[insertColumns (insertLocation： InsertLocation，columnCount： number，values？： string [] [] ) ](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)|使用单元格的列作为模板，在单元格的左侧或右侧添加列。|
||[insertRows (insertLocation： InsertLocation，rowCount： number，values？： string [] [] ) ](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)|使用单元格的行作为模板，在单元格的上方或下方插入行。|
||[body](/javascript/api/word/word.tablecell#body)|获取单元格的 body 对象。|
||[cellIndex](/javascript/api/word/word.tablecell#cellindex)|获取单元格行中的单元格索引。|
||[parentRow](/javascript/api/word/word.tablecell#parentrow)|获取单元格的父行。|
||[parentTable](/javascript/api/word/word.tablecell#parenttable)|获取单元格的父表。|
||[rowIndex](/javascript/api/word/word.tablecell#rowindex)|获取表中单元格行的索引。|
||[width](/javascript/api/word/word.tablecell#width)|获取单元格的宽度（以磅为单位）。|
||[setCellPadding (cellPaddingLocation： CellPaddingLocation，cellPadding： number) ](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)|设置单元格填充（以磅为单位）。|
||[shadingColor](/javascript/api/word/word.tablecell#shadingcolor)|获取或设置单元格的底纹色。|
||[value](/javascript/api/word/word.tablecell#value)|获取并设置单元格的文本。|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)|获取并设置单元格的垂直对齐方式。|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)|获取此集合中的第一个表格单元格。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.tablecellcollection#getfirstornullobject--)|获取此集合中的第一个表格单元格。|
||[items](/javascript/api/word/word.tablecellcollection#items)|获取此集合中已加载的子项。|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getfirst--)|获取此集合中的第一个表格。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.tablecollection#getfirstornullobject--)|获取此集合中的第一个表格。|
||[items](/javascript/api/word/word.tablecollection#items)|获取此集合中已加载的子项。|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear--)|清除行内容。|
||[delete()](/javascript/api/word/word.tablerow#delete--)|删除整行。|
||[getBorder (borderLocation： Word BorderLocation) ](/javascript/api/word/word.tablerow#getborder-borderlocation-)|获取行中单元格的边框样式。|
||[getCellPadding (cellPaddingLocation： Word CellPaddingLocation) ](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)|获取单元格填充（以磅为单位）。|
||[getNext ( # B1 ](/javascript/api/word/word.tablerow#getnext--)|获取下一行。|
||[getNextOrNullObject ( # B1 ](/javascript/api/word/word.tablerow#getnextornullobject--)|获取下一行。|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)|获取并设置行中每个单元格的水平对齐方式。|
||[ignorePunct](/javascript/api/word/word.tablerow#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.tablerow#ignorespace)||
||[insertRows (insertLocation： InsertLocation，rowCount： number，values？： string [] [] ) ](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)|使用此行作为模板插入行。|
||[matchCase](/javascript/api/word/word.tablerow#matchcase)||
||[matchPrefix](/javascript/api/word/word.tablerow#matchprefix)||
||[matchSuffix](/javascript/api/word/word.tablerow#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.tablerow#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.tablerow#matchwildcards)||
||[preferredHeight](/javascript/api/word/word.tablerow#preferredheight)|获取并设置行的首选高度（以磅为单位）。|
||[cellCount](/javascript/api/word/word.tablerow#cellcount)|获取行中的单元格数。|
||[单元](/javascript/api/word/word.tablerow#cells)|获取单元格。|
||[font](/javascript/api/word/word.tablerow#font)|获取字体。|
||[isHeader](/javascript/api/word/word.tablerow#isheader)|检查该行是否为标题行。|
||[parentTable](/javascript/api/word/word.tablerow#parenttable)|获取父表。|
||[rowIndex](/javascript/api/word/word.tablerow#rowindex)|获取其父表中的行索引。|
||[search (searchText： string，searchOptions？： Word. SearchOptions \| {ignorePunct？： Boolean ignoreSpace？： Boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： Boolean matchWholeWord？： Boolean matchWildcards？： boolean} ) ](/javascript/api/word/word.tablerow#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|在行的作用域上使用指定的 SearchOptions 执行搜索。|
||[选择 (selectionMode？： SelectionMode) ](/javascript/api/word/word.tablerow#select-selectionmode-)|选择行，然后将 Word UI 导航到相应位置。|
||[setCellPadding (cellPaddingLocation： CellPaddingLocation，cellPadding： number) ](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)|设置单元格填充（以磅为单位）。|
||[shadingColor](/javascript/api/word/word.tablerow#shadingcolor)|获取并设置底纹色。|
||[values](/javascript/api/word/word.tablerow#values)|获取并设置行中的文本值，作为 2D Javascript 数组。|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)|获取并设置行中单元格的垂直对齐方式。|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)|获取此集合中的第一行。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.tablerowcollection#getfirstornullobject--)|获取此集合中的第一行。|
||[items](/javascript/api/word/word.tablerowcollection#items)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Word JavaScript API 要求集](word-api-requirement-sets.md)

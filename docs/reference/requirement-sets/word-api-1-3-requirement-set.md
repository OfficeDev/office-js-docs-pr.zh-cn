---
title: Word JavaScript API 要求集1。3
description: 有关 WordApi 1.3 要求集的详细信息
ms.date: 07/25/2019
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 6402543ddfb2feaa116de40982dcb61c30c8597b
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430497"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Word JavaScript API 1.3 的最近更新

WordApi 1.3 增加了对内容控件、自定义 XML 和文档级设置的更多支持。

## <a name="api-list"></a>API 列表

下表列出了 Word JavaScript API 要求集1.3 中的 Api。 若要查看 Word JavaScript API 要求集1.3 或更早版本支持的所有 Api 的 API 参考文档，请参阅 [要求集1.3 或更早版本中的 Word api](/javascript/api/word?view=word-js-1.3&preserve-view=true)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument (base64File？： string) ](/javascript/api/word/word.application#createdocument-base64file-)|使用可选的 base64 编码的 .docx 文件创建一个新文档。|
|[正文](/javascript/api/word/word.body)|[getRange (rangeLocation？： Word RangeLocation) ](/javascript/api/word/word.body#getrange-rangelocation-)|获取整个正文或正文的起点/终点，作为一个范围。|
||[insertTable (rowCount： number，columnCount： number，insertLocation： InsertLocation，values？： string [] [] ) ](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|插入包含指定行数和列数的 table。 insertLocation 的可取值为“Start”或“End”。|
||[lists](/javascript/api/word/word.body#lists)|获取 body 中的一组 list 对象。 只读。|
||[parentBody](/javascript/api/word/word.body#parentbody)|获取 body 的父正文。例如，表格单元格 body 的父正文可能是标题。如果不存在父正文控件，则引发。只读。|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|获取 body 的父正文。例如，表格单元格 body 的父正文可能是标题。如果不存在父正文控件，则返回一个 Null 对象。只读。|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|获取包含正文的内容控件。 如果没有父内容控件，则返回 null 对象。 只读。|
||[parentSection](/javascript/api/word/word.body#parentsection)|获取 body 的父节。 如果没有父节，则引发。 只读。|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|获取 body 的父节。 如果没有父节，则返回 null 对象。 只读。|
||[表](/javascript/api/word/word.body#tables)|获取 body 中的一组 table 对象。 只读。|
||[type](/javascript/api/word/word.body#type)|获取 body 的类型。 类型可取值为“MainDoc”、“Section”、“Header”、“Footer”或“TableCell”。 只读。|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|获取或设置 body 的嵌入样式名称。请对可以在区域设置之间移植的嵌入样式使用此属性。若要使用自定义样式或本地化样式名称，请参阅“style”属性。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange (rangeLocation？： Word RangeLocation) ](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|获取整个内容控件或内容控件的起点/终点，作为一个范围。|
||[getTextRanges (endingMarks： string []，trimSpacing？： boolean) ](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|使用标点符号和/或其他结束标记获取内容控件中的文本范围。|
||[insertTable (rowCount： number，columnCount： number，insertLocation： InsertLocation，values？： string [] [] ) ](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|将包含指定行数和列数的 table 插入 contentControl 中或在其旁边插入。 InsertLocation 值可以是 "Start"、"End"、"Before" 或 "After"。|
||[lists](/javascript/api/word/word.contentcontrol#lists)|获取 contentControl 中的一组 list 对象。 只读。|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|获取 contentControl 的父正文。 只读。|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|获取包含此内容控件的内容控件。 如果没有父内容控件，则返回 null 对象。 只读。|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|获取包含 contentControl 的 table。 如果表中不包含此项，则引发。 只读。|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|获取包含 contentControl 的 tableCell。 如果表格单元格中不包含此项，则会引发此异常。 只读。|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|获取包含 contentControl 的 tableCell。 如果未包含在 tableCell 中，则此关系返回空对象。 只读。|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parenttableornullobject)|获取包含 contentControl 的 table。 如果未包含在 table 中，则此关系返回空对象。 只读。|
||[个子](/javascript/api/word/word.contentcontrol#subtype)|获取 contentControl 的子类型。 对于 RTF 格式 contentControl，子类型的可取值为“RichTextInline”、“RichTextParagraphs”、“RichTextTableCell”、“RichTextTableRow”和“RichTextTable”。 只读。|
||[表](/javascript/api/word/word.contentcontrol#tables)|获取 contentControl 中的一组 table 对象。 只读。|
||[拆分 (定界符： string []，multiParagraphs？： boolean，trimDelimiters？： boolean，trimSpacing？： boolean) ](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|使用分隔符将内容控件拆分为各个子范围。|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|获取或设置 contentControl 的嵌入样式名称。请对可以在区域设置之间移植的嵌入样式使用此属性。若要使用自定义样式或本地化样式名称，请参阅“style”属性。|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject (id： number) ](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|按其标识符获取内容控件。 如果此集合中没有包含标识符的内容控件，则返回一个 null 对象。|
||[getByTypes (类型： ContentControlType [] ) ](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|获取具有指定类型和/或子类型的内容控件。|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|获取此集合中的第一个内容控件。 如果此集合为空，则引发。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|获取此集合中的第一个内容控件。 如果此集合为空，则返回 null 对象。|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|删除 custom property 对象。|
||[key](/javascript/api/word/word.customproperty#key)|获取 customProperty 的键。 只读。|
||[type](/javascript/api/word/word.customproperty#type)|获取自定义属性的值类型。 可能的值包括： String、Number、Date、Boolean。 只读。|
||[value](/javascript/api/word/word.customproperty#value)|获取或设置自定义属性的值。 请注意，即使在 web 上的 Word 和 .docx 文件格式允许这些属性任意长，Word 的桌面版本也会将字符串值截断为 255 16 位字符 (可能通过细分代理项对) 来创建无效的 unicode。|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add (key： string，value： any) ](/javascript/api/word/word.custompropertycollection#add-key--value-)|新建自定义属性或设置现有自定义属性。|
||[deleteAll ( # B1 ](/javascript/api/word/word.custompropertycollection#deleteall--)|删除此集合中的所有自定义属性。|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|获取自定义属性的计数。|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|按键获取自定义属性对象（不区分大小写）。 如果自定义属性不存在，则引发此异常。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|按键获取自定义属性对象（不区分大小写）。 如果自定义属性不存在，则返回 null 对象。|
||[items](/javascript/api/word/word.custompropertycollection#items)|获取此集合中已加载的子项。|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|获取文档的属性。 只读。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[打开 ( # B1 ](/javascript/api/word/word.documentcreated#open--)|打开文档。|
||[body](/javascript/api/word/word.documentcreated#body)|获取文档的正文对象。 正文是不包括标头、页脚、脚注、文本框等的文本。 只读。|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|获取文档中的内容控件对象的集合。 这包括文档正文、标头、页脚、文本框等中的内容控件。 只读。|
||[properties](/javascript/api/word/word.documentcreated#properties)|获取文档的属性。 只读。|
||[保存](/javascript/api/word/word.documentcreated#saved)|指示是否已保存在文档中所做的更改。如果值为 true，表示文档自上次保存以来并未更改。只读。|
||[sections](/javascript/api/word/word.documentcreated#sections)|获取文档中的节对象的集合。 只读。|
||[save()](/javascript/api/word/word.documentcreated#save--)|保存文档。 如果文档以前未保存过，将使用 Word 的默认文件命名约定。|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[编写](/javascript/api/word/word.documentproperties#author)|获取或设置 document 的作者。|
||[类别](/javascript/api/word/word.documentproperties#category)|获取或设置 document 的类别。|
||[comments](/javascript/api/word/word.documentproperties#comments)|获取或设置 document 的注释。|
||[company](/javascript/api/word/word.documentproperties#company)|获取或设置 document 的公司。|
||[format](/javascript/api/word/word.documentproperties#format)|获取或设置 document 的格式。|
||[关键字](/javascript/api/word/word.documentproperties#keywords)|获取或设置 document 的关键字。|
||[manager](/javascript/api/word/word.documentproperties#manager)|获取或设置 document 的管理者。|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|获取 document 的应用程序名称。 只读。|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|获取 document 的创建日期。 只读。|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|获取 document 的一组 customProperty。 只读。|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|获取文档的最后一个作者。 只读。|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastprintdate)|获取 document 的上次打印日期。 只读。|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|获取 document 的上次保存日期。 只读。|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|获取 document 的修订号。 只读。|
||[保护](/javascript/api/word/word.documentproperties#security)|获取 document 的安全性。 只读。|
||[template](/javascript/api/word/word.documentproperties#template)|获取 document 的模板。 只读。|
||[subject](/javascript/api/word/word.documentproperties#subject)|获取或设置 document 的主题。|
||[title](/javascript/api/word/word.documentproperties#title)|获取或设置 document 的标题。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext ( # B1 ](/javascript/api/word/word.inlinepicture#getnext--)|获取下一个嵌入式图像。 如果此嵌入式图像是最后一个，则引发。|
||[getNextOrNullObject ( # B1 ](/javascript/api/word/word.inlinepicture#getnextornullobject--)|获取下一个嵌入式图像。 如果此嵌入式图像是最后一个，则返回 null 对象。|
||[getRange (rangeLocation？： Word RangeLocation) ](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|获取图片或图片的起点/终点，作为一个范围。|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|获取包含嵌入式图像的内容控件。 如果没有父内容控件，则返回 null 对象。 只读。|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|获取包含嵌入式图像的 table。 如果表中不包含此项，则引发。 只读。|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|获取包含嵌入式图像的 tableCell。 如果表格单元格中不包含此项，则会引发此异常。 只读。|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|获取包含嵌入式图像的 tableCell。 如果未包含在 tableCell 中，则此关系返回空对象。 只读。|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parenttableornullobject)|获取包含嵌入式图像的 table。 如果未包含在 table 中，则此关系返回空对象。 只读。|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|获取此集合中的第一个嵌入式图像。 如果此集合为空，则引发。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|获取此集合中的第一个嵌入式图像。 如果此集合为空，则返回 null 对象。|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs (级别： number) ](/javascript/api/word/word.list#getlevelparagraphs-level-)|获取列表中指定级别的段落。|
||[getLevelString (级别： number) ](/javascript/api/word/word.list#getlevelstring-level-)|以字符串形式获取指定级别的项目符号、编号或图片。|
||[insertParagraph (paragraphText： string，insertLocation： InsertLocation) ](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|在指定位置插入段落。 InsertLocation 值可以是 "Start"、"End"、"Before" 或 "After"。|
||[id](/javascript/api/word/word.list#id)|获取列表的 id。|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|检查 list 中是否包含所有 9 个级别。值为 true 表示级别存在，即各个级别至少存在一个列表项。只读。|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|获取 list 中的所有 9 个级别类型。 每种类型可以是 "项目符号"、"数字" 或 "图片"。 只读。|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|获取 list 中的段落。 只读。|
||[setLevelAlignment (level： number，对齐方式： Word) 对齐方式 ](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|设置列表中指定级别的项目符号、编号或图片的对齐方式。|
||[setLevelBullet (level： number，listBullet： ListBullet，charCode？： number，fontName？： string) ](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|设置 list 中指定级别的项目符号格式。 如果项目符号为“Custom”，则需要使用字符代码。|
||[setLevelIndents (level： number，textIndent： number，bulletNumberPictureIndent： number) ](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|设置列表中指定级别的两种缩进方式。|
||[setLevelNumbering (level： number，listNumbering： ListNumbering，格式说明符？： Array<string \| number>) ](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|设置列表中指定级别的编号格式。|
||[setLevelStartingNumber (level： number，startingNumber： number) ](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|设置 list 中指定级别的起始编号。 默认值为 1。|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|按标识符获取列表。 如果此集合中没有包含标识符的列表，则会引发此异常。|
||[getByIdOrNullObject (id： number) ](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|按标识符获取列表。 如果此集合中没有包含标识符的列表，则返回一个 null 对象。|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|获取此集合中的第一个列表。 如果此集合为空，则引发。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.listcollection#getfirstornullobject--)|获取此集合中的第一个列表。 如果此集合为空，则返回 null 对象。|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getitem-index-)|按列表对象在集合中的索引获取列表。|
||[items](/javascript/api/word/word.listcollection#items)|获取此集合中已加载的子项。|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor (parentOnly？： boolean) ](/javascript/api/word/word.listitem#getancestor-parentonly-)|获取列表项目的父级或最近的上级元素（如果父级不存在的话）。 如果列表项没有上级，则引发。|
||[getAncestorOrNullObject (parentOnly？： boolean) ](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|获取列表项目的父级或最近的上级元素（如果父级不存在的话）。 如果列表项没有祖先，则返回 null 对象。|
||[getDescendants (directChildrenOnly？： boolean) ](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|获取相应列表项目的所有后代列表项目。|
||[level](/javascript/api/word/word.listitem#level)|获取或设置 list 中项的级别。|
||[listString](/javascript/api/word/word.listitem#liststring)|以字符串形式获取列表项的项目符号、编号或图片。 只读。|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|获取 listItem 相对于同级元素的序号。 只读。|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList (listId： number，level： number) ](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|将 paragraph 加入指定级别的现有 list。如果 paragraph 无法加入 list 或已是 listItem，则无法执行此方法。|
||[detachFromList ( # B1 ](/javascript/api/word/word.paragraph#detachfromlist--)|如果此段落是列表项目的话，从列表中移出此段落。|
||[getNext ( # B1 ](/javascript/api/word/word.paragraph#getnext--)|获取下一个段落。 如果段落是最后一个段落，则引发。|
||[getNextOrNullObject ( # B1 ](/javascript/api/word/word.paragraph#getnextornullobject--)|获取下一个段落。 如果该段落是最后一个，则返回 null 对象。|
||[getPrevious ( # B1 ](/javascript/api/word/word.paragraph#getprevious--)|获取上一个段落。 如果段落是第一个段落，则引发。|
||[getPreviousOrNullObject ( # B1 ](/javascript/api/word/word.paragraph#getpreviousornullobject--)|获取上一个段落。 如果该段落是第一个，则返回 null 对象。|
||[getRange (rangeLocation？： Word RangeLocation) ](/javascript/api/word/word.paragraph#getrange-rangelocation-)|获取整个段落或段落的起点/终点，作为一个范围。|
||[getTextRanges (endingMarks： string []，trimSpacing？： boolean) ](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|使用标点符号和/或其他结束标记获取段落中的文本范围。|
||[insertTable (rowCount： number，columnCount： number，insertLocation： InsertLocation，values？： string [] [] ) ](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|插入包含指定行数和列数的 table。 insertLocation 的可取值为“Before”或“After”。|
||[isLastParagraph](/javascript/api/word/word.paragraph#islastparagraph)|指明 paragraph 是其父正文内的最后一个段落。 只读。|
||[isListItem](/javascript/api/word/word.paragraph#islistitem)|检查 paragraph 是否为 listItem。 只读。|
||[列表](/javascript/api/word/word.paragraph#list)|获取 paragraph 所属的 List。 如果段落不在列表中，将引发此异常。 只读。|
||[listItem](/javascript/api/word/word.paragraph#listitem)|获取 paragraph 的 ListItem。 如果段落不是列表的一部分，则引发。 只读。|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listitemornullobject)|获取 paragraph 的 ListItem。 如果 paragraph 未包含在 list 中，则此关系返回空对象。 只读。|
||[listOrNullObject](/javascript/api/word/word.paragraph#listornullobject)|获取 paragraph 所属的 List。 如果 paragraph 未包含在 list 中，则此关系返回空对象。 只读。|
||[parentBody](/javascript/api/word/word.paragraph#parentbody)|获取 paragraph 的父正文。 只读。|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentcontentcontrolornullobject)|获取包含段落的内容控件。 如果没有父内容控件，则返回 null 对象。 只读。|
||[parentTable](/javascript/api/word/word.paragraph#parenttable)|获取包含 paragraph 的 table。 如果表中不包含此项，则引发。 只读。|
||[parentTableCell](/javascript/api/word/word.paragraph#parenttablecell)|获取包含 paragraph 的 tableCell。 如果表格单元格中不包含此项，则会引发此异常。 只读。|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parenttablecellornullobject)|获取包含 paragraph 的 tableCell。 如果未包含在 tableCell 中，则此关系返回空对象。 只读。|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parenttableornullobject)|获取包含 paragraph 的 table。 如果未包含在 table 中，则此关系返回空对象。 只读。|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tablenestinglevel)|获取 paragraph 的表级别。 如果 paragraph 未包含在 table 中，则此属性返回 0。 只读。|
||[拆分 (定界符： string []，trimDelimiters？： boolean，trimSpacing？： boolean) ](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|使用分隔符将段落拆分为多个子范围。|
||[startNewList ( # B1 ](/javascript/api/word/word.paragraph#startnewlist--)|生成包含此 paragraph 的新 list。 如果此 paragraph 已是 listItem，则无法执行此方法。|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|获取或设置 paragraph 的嵌入样式名称。请对可以在区域设置之间移植的嵌入样式使用此属性。若要使用自定义样式或本地化样式名称，请参阅“style”属性。|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|获取此集合中的第一个段落。 如果集合为空，则引发。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|获取此集合中的第一个段落。 如果集合为空，则返回 null 对象。|
||[getLast ( # B1 ](/javascript/api/word/word.paragraphcollection#getlast--)|获取此集合中的最后一个段落。 如果集合为空，则引发。|
||[getLastOrNullObject ( # B1 ](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|获取此集合中的最后一个段落。 如果集合为空，则返回 null 对象。|
|[区域](/javascript/api/word/word.range)|[compareLocationWith (范围： Word) 区域 ](/javascript/api/word/word.range#comparelocationwith-range-)|比较此范围与另一范围的位置。|
||[expandTo (范围： Word) 区域 ](/javascript/api/word/word.range#expandto-range-)|返回从此 range 进行任一方向扩展的新 range，以便覆盖另一 range。 此 range 不变。 如果两个区域没有 union，则引发。|
||[expandToOrNullObject (范围： Word) 区域 ](/javascript/api/word/word.range#expandtoornullobject-range-)|返回从此 range 进行任一方向扩展的新 range，以便覆盖另一 range。 此 range 不变。 如果两个区域没有 union，则返回 null 对象。|
||[getHyperlinkRanges ( # B1 ](/javascript/api/word/word.range#gethyperlinkranges--)|获取相应范围内的超链接子范围。|
||[getNextTextRange (endingMarks： string []，trimSpacing？： boolean) ](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|使用标点符号和/或其他结束标记获取下一个文本范围。 如果此文本范围是最后一个，则引发。|
||[getNextTextRangeOrNullObject (endingMarks： string []，trimSpacing？： boolean) ](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|使用标点符号和/或其他结束标记获取下一个文本范围。 如果此文本范围是最后一个，则返回 null 对象。|
||[getRange (rangeLocation？： Word RangeLocation) ](/javascript/api/word/word.range#getrange-rangelocation-)|克隆相应范围，或获取该范围的起点/终点作为一个新范围。|
||[getTextRanges (endingMarks： string []，trimSpacing？： boolean) ](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|使用标点符号和/或其他结束标记获取区域中的文本子范围。|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|获取 range 内的第一个超链接，或在 range 内设置超链接。 在 range 内设置新的超链接将删除 range 内的所有超链接。 使用 "#" 将地址部分与可选位置部分分开。|
||[insertTable (rowCount： number，columnCount： number，insertLocation： InsertLocation，values？： string [] [] ) ](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|插入包含指定行数和列数的 table。 insertLocation 的可取值为“Before”或“After”。|
||[intersectWith (范围： Word) 区域 ](/javascript/api/word/word.range#intersectwith-range-)|返回新 range 作为此 range 与另一 range 的交集。 此 range 不变。 如果两个区域不重叠或相邻，则引发。|
||[intersectWithOrNullObject (范围： Word) 区域 ](/javascript/api/word/word.range#intersectwithornullobject-range-)|返回新 range 作为此 range 与另一 range 的交集。 此 range 不变。 如果两个区域不重叠或相邻，则返回 null 对象。|
||[isEmpty](/javascript/api/word/word.range#isempty)|检查 range 长度是否为零。 只读。|
||[lists](/javascript/api/word/word.range#lists)|获取 range 中的一组 list 对象。 只读。|
||[parentBody](/javascript/api/word/word.range#parentbody)|获取 range 的父正文。 只读。|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|获取包含该范围的内容控件。 如果没有父内容控件，则返回 null 对象。 只读。|
||[parentTable](/javascript/api/word/word.range#parenttable)|获取包含 range 的 table。 如果表中不包含此项，则引发。 只读。|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|获取包含 range 的 tableCell。 如果表格单元格中不包含此项，则会引发此异常。 只读。|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parenttablecellornullobject)|获取包含 range 的 tableCell。 如果未包含在 tableCell 中，则此关系返回空对象。 只读。|
||[parentTableOrNullObject](/javascript/api/word/word.range#parenttableornullobject)|获取包含 range 的 table。 如果未包含在 table 中，则此关系返回空对象。 只读。|
||[表](/javascript/api/word/word.range#tables)|获取 range 中的一组 table 对象。 只读。|
||[拆分 (定界符： string []，multiParagraphs？： boolean，trimDelimiters？： boolean，trimSpacing？： boolean) ](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|使用分隔符将相应范围拆分为各个子范围。|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|获取或设置 range 的嵌入样式名称。请对可以在区域设置之间移植的嵌入样式使用此属性。若要使用自定义样式或本地化样式名称，请参阅“style”属性。|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|获取此集合中的第一个范围。 如果此集合为空，则引发。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.rangecollection#getfirstornullobject--)|获取此集合中的第一个范围。 如果此集合为空，则返回 null 对象。|
|[Section](/javascript/api/word/word.section)|[getNext ( # B1 ](/javascript/api/word/word.section#getnext--)|获取下一节。 如果此部分是最后一个，则引发。|
||[getNextOrNullObject ( # B1 ](/javascript/api/word/word.section#getnextornullobject--)|获取下一节。 如果此节是最后一个，则返回 null 对象。|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|获取此集合中的第一节。 如果此集合为空，则引发。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|获取此集合中的第一节。 如果此集合为空，则返回 null 对象。|
|[Table](/javascript/api/word/word.table)|[addColumns (insertLocation： InsertLocation，columnCount： number，values？： string [] [] ) ](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|使用第一个或最后一个现有列作为模板，将列添加到 table 的开头或结尾。此方法适用于一致的 table。字符串值（若指定）是在新插入的行中进行设置。|
||[addRows (insertLocation： InsertLocation，rowCount： number，values？： string [] [] ) ](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|使用第一个或最后一个现有行作为模板，将行添加到 table 的开头或结尾。字符串值（若指定）是在新插入的行中进行设置。|
||[对齐方式](/javascript/api/word/word.table#alignment)|获取或设置表相对于页列的对齐方式。 该值可以是 "Left"、"居中" 或 "Right"。|
||[autoFitWindow ( # B1 ](/javascript/api/word/word.table#autofitwindow--)|自动调整表列，以适应窗口的宽度。|
||[clear()](/javascript/api/word/word.table#clear--)|清除表内容。|
||[delete()](/javascript/api/word/word.table#delete--)|删除整个表格。|
||[deleteColumns (columnIndex： number，columnCount？： number) ](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|删除特定的列。 此方法适用于一致的 table。|
||[deleteRows (rowIndex： number，rowCount？： number) ](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|删除特定的行。|
||[distributeColumns ( # B1 ](/javascript/api/word/word.table#distributecolumns--)|将列设置为等宽。 此方法适用于一致的 table。|
||[getBorder (borderLocation： Word BorderLocation) ](/javascript/api/word/word.table#getborder-borderlocation-)|获取指定边框的边框样式。|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|获取指定行和列处的表单元格。 如果指定的表格单元格不存在，则引发。|
||[getCellOrNullObject (rowIndex： number，cellIndex： number) ](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|获取指定行和列处的表单元格。 如果指定的表格单元格不存在，则返回 null 对象。|
||[getCellPadding (cellPaddingLocation： Word CellPaddingLocation) ](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|获取单元格填充（以磅为单位）。|
||[getNext ( # B1 ](/javascript/api/word/word.table#getnext--)|获取下一个表格。 如果此表是最后一个表，则会引发此异常。|
||[getNextOrNullObject ( # B1 ](/javascript/api/word/word.table#getnextornullobject--)|获取下一个表格。 如果此表是最后一个，则返回 null 对象。|
||[getParagraphAfter ( # B1 ](/javascript/api/word/word.table#getparagraphafter--)|获取 table 之后的 paragraph。 如果表后面没有段落，则会引发此异常。|
||[getParagraphAfterOrNullObject ( # B1 ](/javascript/api/word/word.table#getparagraphafterornullobject--)|获取 table 之后的 paragraph。 如果表后面没有段落，则返回一个 null 对象。|
||[getParagraphBefore ( # B1 ](/javascript/api/word/word.table#getparagraphbefore--)|获取 table 之前的 paragraph。 如果表之前没有段落，则引发此异常。|
||[getParagraphBeforeOrNullObject ( # B1 ](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|获取 table 之前的 paragraph。 如果表之前没有段落，则返回一个 null 对象。|
||[getRange (rangeLocation？： Word RangeLocation) ](/javascript/api/word/word.table#getrange-rangelocation-)|获取包含此表格的范围，或包含此表格的开头或结尾的范围。|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|获取并设置标题行数。|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|获取并设置 table 中每个单元格的水平对齐方式。 该值可以是 "Left"、"居中"、"Right" 或 "两端对齐"。|
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|在表格中插入内容控件。|
||[insertParagraph (paragraphText： string，insertLocation： InsertLocation) ](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|在指定位置插入段落。 insertLocation 值可以为“Before”或“After”。|
||[insertTable (rowCount： number，columnCount： number，insertLocation： InsertLocation，values？： string [] [] ) ](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|插入包含指定行数和列数的 table。 insertLocation 的可取值为“Before”或“After”。|
||[font](/javascript/api/word/word.table#font)|获取字体。 使用此关系可获取并设置字体名称、大小、颜色和其他属性。 只读。|
||[isUniform](/javascript/api/word/word.table#isuniform)|指明所有表行是否一致。 只读。|
||[nestingLevel](/javascript/api/word/word.table#nestinglevel)|获取 table 的嵌套级别。 顶级 table 的级别为 1。 只读。|
||[parentBody](/javascript/api/word/word.table#parentbody)|获取 table 的父正文。 只读。|
||[parentContentControl](/javascript/api/word/word.table#parentcontentcontrol)|获取包含 table 的 contentControl。 如果没有父内容控件，将引发此异常。 只读。|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentcontentcontrolornullobject)|获取包含 table 的 contentControl。 如果没有父内容控件，则返回 null 对象。 只读。|
||[parentTable](/javascript/api/word/word.table#parenttable)|获取包含此 table 的 table。 如果表中不包含此项，则引发。 只读。|
||[parentTableCell](/javascript/api/word/word.table#parenttablecell)|获取包含此 table 的 tableCell。 如果表格单元格中不包含此项，则会引发此异常。 只读。|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parenttablecellornullobject)|获取包含此 table 的 tableCell。 如果未包含在 tableCell 中，则此关系返回空对象。 只读。|
||[parentTableOrNullObject](/javascript/api/word/word.table#parenttableornullobject)|获取包含此 table 的 table。 如果未包含在 table 中，则此关系返回空对象。 只读。|
||[rowCount](/javascript/api/word/word.table#rowcount)|获取表格中的行数。 只读。|
||[rows](/javascript/api/word/word.table#rows)|获取所有表格行。 只读。|
||[表](/javascript/api/word/word.table#tables)|获取嵌套一级的子 table。 只读。|
||[search (searchText： string，searchOptions？： Word. SearchOptions](/javascript/api/word/word.table#search-searchtext--searchoptions-)|在表对象的作用域上使用指定的 SearchOptions 执行搜索。 搜索结果是 range 对象的集合。|
||[选择 (selectionMode？： SelectionMode) ](/javascript/api/word/word.table#select-selectionmode-)|选择表格或其开头或结尾位置，然后将 Word UI 导航到相应位置。|
||[setCellPadding (cellPaddingLocation： CellPaddingLocation，cellPadding： number) ](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)|设置单元格填充（以磅为单位）。|
||[shadingColor](/javascript/api/word/word.table#shadingcolor)|获取并设置底纹色。 按“#RRGGBB”格式或使用颜色名称指定颜色。|
||[style](/javascript/api/word/word.table#style)|获取或设置 table 的样式名称。请对自定义样式和本地化样式名称使用此属性。若要使用可以在区域设置之间移植的嵌入样式，请参阅“styleBuiltIn”属性。|
||[styleBandedColumns](/javascript/api/word/word.table#stylebandedcolumns)|获取并设置 table 是否有镶边列。|
||[styleBandedRows](/javascript/api/word/word.table#stylebandedrows)|获取并设置 table 是否有镶边行。|
||[styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)|获取或设置 table 的嵌入样式名称。请对可以在区域设置之间移植的嵌入样式使用此属性。若要使用自定义样式或本地化样式名称，请参阅“style”属性。|
||[styleFirstColumn](/javascript/api/word/word.table#stylefirstcolumn)|获取并设置 table 的第一列是否采用特殊样式。|
||[styleLastColumn](/javascript/api/word/word.table#stylelastcolumn)|获取并设置 table 的最后一列是否采用特殊样式。|
||[styleTotalRow](/javascript/api/word/word.table#styletotalrow)|获取并设置 table 的总计行（最后一行）是否采用特殊样式。|
||[values](/javascript/api/word/word.table#values)|以 2D Javascript 数组形式获取并设置 table 中的文本值。|
||[verticalAlignment](/javascript/api/word/word.table#verticalalignment)|获取并设置 table 中每个单元格的垂直对齐方式。 值可以是 ' Top '、' Center ' 或 ' 底端 '。|
||[width](/javascript/api/word/word.table#width)|获取并设置 table 的宽度（以磅为单位）。|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|获取或设置表格边框颜色。|
||[type](/javascript/api/word/word.tableborder#type)|获取或设置表边框的类型。|
||[width](/javascript/api/word/word.tableborder#width)|获取或设置表边框的宽度（以磅为单位）。 不适用于宽度固定的表边框类型。|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnwidth)|获取并设置单元格的列宽度（以磅为单位）。 此方法适用于一致的 table。|
||[deleteColumn ( # B1 ](/javascript/api/word/word.tablecell#deletecolumn--)|删除包含该单元格的列。 此方法适用于一致的 table。|
||[deleteRow ( # B1 ](/javascript/api/word/word.tablecell#deleterow--)|删除包含该单元格的行。|
||[getBorder (borderLocation： Word BorderLocation) ](/javascript/api/word/word.tablecell#getborder-borderlocation-)|获取指定边框的边框样式。|
||[getCellPadding (cellPaddingLocation： Word CellPaddingLocation) ](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)|获取单元格填充（以磅为单位）。|
||[getNext ( # B1 ](/javascript/api/word/word.tablecell#getnext--)|获取下一个单元格。 如果此单元格是最后一个单元格，则引发。|
||[getNextOrNullObject ( # B1 ](/javascript/api/word/word.tablecell#getnextornullobject--)|获取下一个单元格。 如果此单元格是最后一个，则返回 null 对象。|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)|获取并设置单元格的水平对齐方式。 该值可以是 "Left"、"居中"、"Right" 或 "两端对齐"。|
||[insertColumns (insertLocation： InsertLocation，columnCount： number，values？： string [] [] ) ](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)|使用单元格的列作为模板，在单元格的左侧或右侧添加列。此方法适用于一致的 table。字符串值（若指定）是在新插入的行中进行设置。|
||[insertRows (insertLocation： InsertLocation，rowCount： number，values？： string [] [] ) ](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)|使用单元格的行作为模板，在单元格的上方或下方插入行。字符串值（若指定）是在新插入的行中进行设置。|
||[body](/javascript/api/word/word.tablecell#body)|获取单元格的 body 对象。 只读。|
||[cellIndex](/javascript/api/word/word.tablecell#cellindex)|获取单元格行中的单元格索引。 只读。|
||[parentRow](/javascript/api/word/word.tablecell#parentrow)|获取单元格的父行。 只读。|
||[parentTable](/javascript/api/word/word.tablecell#parenttable)|获取单元格的父表。 只读。|
||[rowIndex](/javascript/api/word/word.tablecell#rowindex)|获取表中单元格行的索引。 只读。|
||[width](/javascript/api/word/word.tablecell#width)|获取单元格的宽度（以磅为单位）。 只读。|
||[setCellPadding (cellPaddingLocation： CellPaddingLocation，cellPadding： number) ](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)|设置单元格填充（以磅为单位）。|
||[shadingColor](/javascript/api/word/word.tablecell#shadingcolor)|获取或设置单元格的底纹色。 按“#RRGGBB”格式或使用颜色名称指定颜色。|
||[value](/javascript/api/word/word.tablecell#value)|获取并设置单元格的文本。|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)|获取并设置单元格的垂直对齐方式。 值可以是 ' Top '、' Center ' 或 ' 底端 '。|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)|获取此集合中的第一个表格单元格。 如果此集合为空，则引发。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.tablecellcollection#getfirstornullobject--)|获取此集合中的第一个表格单元格。 如果此集合为空，则返回 null 对象。|
||[items](/javascript/api/word/word.tablecellcollection#items)|获取此集合中已加载的子项。|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getfirst--)|获取此集合中的第一个表格。 如果此集合为空，则引发。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.tablecollection#getfirstornullobject--)|获取此集合中的第一个表格。 如果此集合为空，则返回 null 对象。|
||[items](/javascript/api/word/word.tablecollection#items)|获取此集合中已加载的子项。|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear--)|清除行内容。|
||[delete()](/javascript/api/word/word.tablerow#delete--)|删除整行。|
||[getBorder (borderLocation： Word BorderLocation) ](/javascript/api/word/word.tablerow#getborder-borderlocation-)|获取行中单元格的边框样式。|
||[getCellPadding (cellPaddingLocation： Word CellPaddingLocation) ](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)|获取单元格填充（以磅为单位）。|
||[getNext ( # B1 ](/javascript/api/word/word.tablerow#getnext--)|获取下一行。 如果此行是最后一个，则引发。|
||[getNextOrNullObject ( # B1 ](/javascript/api/word/word.tablerow#getnextornullobject--)|获取下一行。 如果此行是最后一个，则返回一个 null 对象。|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)|获取并设置行中每个单元格的水平对齐方式。 该值可以是 "Left"、"居中"、"Right" 或 "两端对齐"。|
||[insertRows (insertLocation： InsertLocation，rowCount： number，values？： string [] [] ) ](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)|使用此行作为模板插入行。 如果值已指定，请将值插入新行。|
||[preferredHeight](/javascript/api/word/word.tablerow#preferredheight)|获取并设置行的首选高度（以磅为单位）。|
||[cellCount](/javascript/api/word/word.tablerow#cellcount)|获取行中的单元格数。 只读。|
||[单元](/javascript/api/word/word.tablerow#cells)|获取单元格。 只读。|
||[font](/javascript/api/word/word.tablerow#font)|获取字体。 使用此关系可获取并设置字体名称、大小、颜色和其他属性。 只读。|
||[isHeader](/javascript/api/word/word.tablerow#isheader)|检查该行是否为标题行。 只读。 若要设置标题行数，请对 Table 对象使用 HeaderRowCount。|
||[parentTable](/javascript/api/word/word.tablerow#parenttable)|获取父表。 只读。|
||[rowIndex](/javascript/api/word/word.tablerow#rowindex)|获取其父表中的行索引。 只读。|
||[search (searchText： string，searchOptions？： Word SearchOptions) ](/javascript/api/word/word.tablerow#search-searchtext--searchoptions-)|在行的作用域上使用指定的 SearchOptions 执行搜索。 搜索结果是 range 对象的集合。|
||[选择 (selectionMode？： SelectionMode) ](/javascript/api/word/word.tablerow#select-selectionmode-)|选择行，然后将 Word UI 导航到相应位置。|
||[setCellPadding (cellPaddingLocation： CellPaddingLocation，cellPadding： number) ](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)|设置单元格填充（以磅为单位）。|
||[shadingColor](/javascript/api/word/word.tablerow#shadingcolor)|获取并设置底纹色。 按“#RRGGBB”格式或使用颜色名称指定颜色。|
||[values](/javascript/api/word/word.tablerow#values)|获取并设置行中的文本值，作为 2D Javascript 数组。|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)|获取并设置行中单元格的垂直对齐方式。 值可以是 ' Top '、' Center ' 或 ' 底端 '。|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)|获取此集合中的第一行。 如果此集合为空，则引发。|
||[getFirstOrNullObject ( # B1 ](/javascript/api/word/word.tablerowcollection#getfirstornullobject--)|获取此集合中的第一行。 如果此集合为空，则返回 null 对象。|
||[items](/javascript/api/word/word.tablerowcollection#items)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Word JavaScript API 要求集](word-api-requirement-sets.md)

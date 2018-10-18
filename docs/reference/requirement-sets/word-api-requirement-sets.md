# <a name="word-javascript-api-requirement-sets"></a>Word JavaScript API 要求集

要求集是 API 成员的具名组 。 Office 加载项使用清单中指定要求集或使用运行时检查，以确定 Office 主机是否支持加载项所需的 API。 有关更多信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

Word 加载项运行跨多个版本的 Office，包括 Windows 版 Office 2016 或更高版本 、Office for iPad、 Office for Mac 和 Office Online。 下表列出了 Word 要求集、支持这些要求集的 Office 宿主应用程序以及应用程序的内部版本或版本号码。

> [!NOTE]
> 对于标记为 Beta 的要求集，请使用指定（或更高）版本的 Office 软件并使用 CDN 上的 Beta 库： https://appsforoffice.microsoft.com/lib/beta/hosted/office.js 。
> 
> 未列出为 Beta 的条目通常可用，您可以继续使用生产 CDN 库： https://appsforoffice.microsoft.com/lib/1/hosted/office.js

|  要求集  |   Office 365 for Windows\*  |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  | Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| WordApi 1.3 | 版本 1612（内部版本 7668.1000）或更高版本| 2017 年 3 月，2.22 或更高版本 | 2017 年 3 月，15.32 或更高版本| 2017 年 3 月 ||
| WordApi 1.2  | 2015 年 12 月更新，版本 1601（内部版本 6568.1000）或更高版本 | 2016 年 1 月，1.18 或更高版本 | 2016 年 1 月，15.19 或更高版本| 2016 年 9 月 | |
| WordApi 1.1  | 版本 1509（内部版本 4266.1001）或更高版本| 2016 年 1 月，1.18 或更高版本 | 2016 年 1 月，15.19 或更高版本| 2016 年 9 月 | |

> [!NOTE]
> 通过 MSI 安装的 Office 2016 的内部版本号为 16.0.4266.1001。 此版本中仅包含 WordApi 1.1 要求集。

要详细了解版本、内部版本号和 Office Online Server，请参阅：

- [更新频道发布的 Office 365 客户端版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概述](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 公用 API 要求集

有关公用 API 要求集的信息，请参阅 [Office 公用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="whats-new-in-word-javascript-api-13"></a>Word JavaScript API 1.3 的最近更新 

下面介绍了要求集 1.3 中 Word JavaScript API 的新增内容。 

|对象| 最近更新| 说明|要求集| 
|:-----|-----|:----|:----| 
|[application](/javascript/api/word/word.application)|_方法_ > createDocument(base64File: string) | 使用 base64 编码的.docx 文件新建文档。 只读。|1.3|
|[body](/javascript/api/word/word.body)|_关系_ > lists|获取 body 中的一组 list 对象。只读。|1.3|
|[body](/javascript/api/word/word.body)|_关系_ > parentBody|获取 body 的父正文。例如，表格单元格 body 的父正文可能是标题。只读。|1.3|
|[body](/javascript/api/word/word.body)|_关系_ > parentSection|获取 body 的父节。只读。|1.3|
|[body](/javascript/api/word/word.body)|_关系_ > styleBuiltIn|获取或设置 body 的嵌入样式名称。请对可以在区域设置之间移植的嵌入样式使用此属性。若要使用自定义样式或本地化样式名称，请参阅 "style" 属性。|1.3|
|[body](/javascript/api/word/word.body)|_关系_ > tables|获取 body 中的一组 table 对象。只读。|1.3|
|[body](/javascript/api/word/word.body)|_关系_ >type|获取 body 的类型。类型可取值为 'MainDoc'、'Section'、'Header'、'Footer' 或 'TableCell'。只读。|1.3|
|[body](/javascript/api/word/word.body)|_方法_ > getRange(rangeLocation: RangeLocation)|获取整个 body 或 body 的起磅/终磅作为范围。|1.3|
|[body](/javascript/api/word/word.body)|_方法_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|插入包含指定行数和列数的表格。insertLocation 的可取值为 'Start' 或 'End'。|1.3|
|[breaktype](/javascript/api/word/word.breaktype)|_关系_ > breaks|指定中断的形式：行、页或节类型。 只读。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_关系_ > lists|获取内容控件中的一组 list 对象。只读。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_关系_ > parentBody|获取内容控件的父正文。只读。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_关系_ > parentTable|获取包含内容控件的表格。如果未包含在表格中，则此关系返回空对象。只读。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_关系_ > parentTableCell|获取包含内容控件的表格单元格。如果未包含在表格单元格中，则此关系返回空对象。只读。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_关系_ > styleBuiltIn|获取或设置内容控件的嵌入样式名称。请对可以在区域设置之间移植的嵌入样式使用此属性。若要使用自定义样式或本地化样式名称，请参阅 "style" 属性。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_关系_ > subtype|获取内容控件的子类型。对于 RTF 格式内容控件，子类型的可取值为 'RichTextInline'、'RichTextParagraphs'、'RichTextTableCell'、'RichTextTableRow'和 'RichTextTable'。只读。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_关系_ > tables|获取内容控件中的一组 table 对象。只读。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_方法_ > getRange(rangeLocation: RangeLocation)|获取整个内容控件或其起磅/终磅作为范围。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_方法_ > getTextRanges(endingMarks: string, trimSpacing: bool)|使用标点符号和/或其他结束标记获取内容控件中的文本区域。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_方法_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|将包含指定行数和列数的表格插入内容控件中或在其旁边插入。insertLocation 的可取值为 'Start'、'End'、'Before' 或 'After'。|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_方法_ > split(delimiters: string, multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|使用分隔符将内容控件拆分为多个子范围。|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_方法_ > getByTypes(types: ContentControlType)|获取具有指定类型和/或子类型的内容控件。|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_方法_ > getFirst()|获取此集合中的第一个内容控件。|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_属性_ > key|获取自定义属性的键。 只读。 |1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_属性_ > value|获取或设置自定义属性的值。|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_关系_ >type|获取自定义属性的值。 只读。|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_方法_ > delete()|删除自定义属性。|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_属性_ > items|一组 CustomProperty 对象。只读。|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_方法_ > deleteAll()|删除此集合中的所有自定义属性。|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_方法_ > getCount()|获取自定义属性的计数。|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_方法_ > getItem(key: string)|按键获取自定义属性对象（不区分大小写）。|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_方法_ > set(key: string, value: object)|创建或设置自定义属性。|1.3|
|[document](/javascript/api/word/word.document)|_关系_ > properties|获取当前文档的属性。只读。|1.3|
|[document](/javascript/api/word/word.document)|_方法_ > open()|打开文档。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > applicationName|获取文档的应用程序名称。 只读。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > author|获取或设置文档的作者。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > category|获取或设置文档的类别。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > comments|获取或设置文档的注释。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > company|获取或设置文档的公司。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > format|获取或设置文档的格式。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > keywords|获取或设置文档的关键字。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > lastAuthor|获取或设置文档的上一作者。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > manager|获取或设置文档的管理者。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > revisionNumber|获取文档的修订号。 只读。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > security|获取文档的安全性。 只读。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > subject|获取或设置文档的主题。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > template|获取文档的模板。 只读。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_属性_ > title|获取或设置文档的标题。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_关系_ > creationDate|获取文档的创建日期。 只读。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_关系_ > customProperties|获取文档的自定义属性的集合。只读。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_关系_ > lastPrintDate|获取文档的上次打印日期。 只读。|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_关系_ > lastSaveTime|获取上一次保存文档的时间。 只读。|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_关系_ > parentTable|获取包含内联图像的表格。如果未包含在表格中，则此关系返回空对象。只读。|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_关系_ > parentTableCell|获取包含嵌入式图像的 tableCell。如果未包含在 tableCell 中，则此关系返回空对象。只读。|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > getNext()|获取下一个内联图像。|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > getRange(rangeLocation: RangeLocation)|获取图片或图片的起磅/终磅作为范围。|1.3|
|[inlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|_方法_ > getFirst()|获取此集合中的第一个内联图像。|1.3|
|[list](/javascript/api/word/word.list)|属性 > id_ _|获取列表的 ID。只读。|1.3|
|[list](/javascript/api/word/word.list)|_属性_ > levelExistences|检查列表中是否包含所有 9 个级别。值为 true 表示级别存在，即各个级别至少存在一个列表项。只读。|1.3|
|[list](/javascript/api/word/word.list)|_关系_ > levelTypes|获取列表中的所有 9 个级别类型。每种类型的可取值为 'Bullet'、'Number' 或 'Picture'。只读。|1.3|
|[list](/javascript/api/word/word.list)|_关系_ > paragraphs|获取列表中的段落。只读。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > getLevelParagraphs(level: number)|获取列表中指定级别的段落。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > getLevelString(level: number)|以字符串形式获取指定级别的项目符号、编号或图片。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > insertParagraph(paragraphText: string, insertLocation: InsertLocation)|在指定位置插入段落。insertLocation 的可取值为 'Start'、'End”、'Before' 或 'After'。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > setLevelAlignment(level: number, alignment: Alignment)|设置列表中指定级别的项目符号、编号或图片的对齐方式。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > setLevelBullet(level: number, listBullet: ListBullet, charCode: number, fontName: string)|设置列表中指定级别的项目符号格式。如果项目符号为 'Custom'，则需要 charCode。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > setLevelIndents(level: number, textIndent: float, textIndent: float)|设置列表中指定级别的两种缩进方式。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > setLevelNumbering(level: number, listNumbering: ListNumbering, formatString: object)|设置列表中指定级别的编号格式。|1.3|
|[list](/javascript/api/word/word.list)|_方法_ > setLevelStartingNumber(level: number, startingNumber: number)|设置列表中指定级别的起始编号。默认值为 1。|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_属性_ > items|一组 list 对象。只读。|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_方法_ > getById(id: number)|按标识符获取列表。|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_方法_ > getFirst()|获取此集合中的第一个列表。|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_方法_ > getItem(index: number)|按 list 对象在集合中的索引获取此对象。|1.3|
|[listItem](/javascript/api/word/word.listitem)|_属性_ > level|获取或设置列表中项的级别。|1.3|
|[listItem](/javascript/api/word/word.listitem)|_属性_ > listString|以字符串形式获取列表项目的项目符号、编号或图片。只读。|1.3|
|[listItem](/javascript/api/word/word.listitem)|_属性_ > siblingIndex|获取列表项相对于同级元素的序号。只读。|1.3|
|[listItem](/javascript/api/word/word.listitem)|_方法_ > getAncestor(parentOnly: bool)|获取父级列表项或最近的上级元素（如果父级不存在的话）。|1.3|
|[listItem](/javascript/api/word/word.listitem)|_方法_ > getDescendants(directChildrenOnly: bool)|获取列表项的所有子级。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|属性 > isLastParagraph_ _|指明段落是其父正文内的最后一个段落。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_属性_ > isListItem|检查段落是否为列表项。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_属性_ > tableNestingLevel|获取段落表格的级别。如果段落未包含在表格中，则此属性返回 0。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_关系_ > list|获取段落所属的列表。如果段落未包含在列表中，则此关系返回空对象。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_关系_ > listItem|获取段落的列表项。如果段落未包含在列表中，则返回空对象。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_关系_ > parentBody|获取段落的父正文。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_关系_ > parentTable|获取包含段落的表格。如果未包含在表格中，则返回空对象。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_关系_ > parentTableCell|获取包含段落的表格单元格。如果未包含在单元格中，则返回空对象。只读。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_关系_ > styleBuiltIn|获取或设置段落的嵌入样式名称。请对可以在区域设置之间移植的嵌入样式使用此属性。若要使用自定义样式或本地化样式名称，请参阅 "style" 属性。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > attachToList(listId: number, level: number)|将段落加入指定级别的现有列表。如果段落无法加入列表或已经是列表项，则执行失败。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > detachFromList()|从列表中移出此段落（如果段落是列表项）。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > getNext()|获取下一个段落。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > getPrevious()|获取上一个段落。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > getRange(rangeLocation: RangeLocation)|获取整个段落或正文的起磅/终磅作为范围。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > getTextRanges(endingMarks: string, trimSpacing: bool)|使用标点符号和/或其他结束标记获取段落中的文本区域。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|插入包含指定行数和列数的表格。insertLocation 的可取值为 'Before' 或 'After'。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > split(delimiters: string, multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|使用分隔符将段落拆分为各个子区域。|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_方法_ > startNewList()|生成包含此段落的新列表。如果此段落已是列表，则无法执行此方法。|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_方法_ > getFirst()|获取此集合中的第一个段落。|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_方法_ > getLast()|获取此集合中的最后一个段落。|1.3|
|[range](/javascript/api/word/word.range)|_属性_ > hyperlink|获取范围内的第一个超链接，或在范围内设置超链接。在范围内设置新的超链接将删除范围内的所有超链接。请使用换行符 ('\n') 隔开地址部分和可选的位置部分。|1.3|
|[range](/javascript/api/word/word.range)|_属性_ > isEmpty|检查范围长度是否为零。只读。|1.3|
|[range](/javascript/api/word/word.range)|_关系_ > lists|获取范围中的一组 list 对象。只读。|1.3|
|[range](/javascript/api/word/word.range)|_关系_ > parentBody|获取范围的父正文。只读。|1.3|
|[range](/javascript/api/word/word.range)|_关系_ > parentTable|获取包含范围的表格。如果未包含在表格中，则返回空。只读。|1.3|
|[range](/javascript/api/word/word.range)|_关系_ > parentTableCell|获取包含范围的表格单元格。如果未包含在表格单元格中，则返回空对象。只读。|1.3|
|[range](/javascript/api/word/word.range)|_关系_ > styleBuiltIn|获取或设置范围的嵌入样式名称。请对可以在区域设置之间移植的嵌入样式使用此属性。若要使用自定义样式或本地化样式名称，请参阅 "style" 属性。|1.3|
|[range](/javascript/api/word/word.range)|_关系_ > tables|获取范围中的一组 table 对象。只读。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > compareLocationWith(range: Range)|比较此范围与另一范围的位置。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > expandTo(range: Range)|返回从此范围进行任一方向扩展的新范围，以便覆盖另一范围。此范围不变。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > getHyperlinkRanges()|获取范围内的超链接子范围。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > getNextTextRange(endingMarks: string, trimSpacing: bool)|使用标点符号和/或其他结束标记获取下一个文本区域。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > getRange(rangeLocation: RangeLocation)|克隆范围，或获取范围的起磅/终磅作为新范围。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > getTextRanges(endingMarks: string, trimSpacing: bool)|使用标点符号和/或其他结束标记获取范围中的子文本范围。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|插入包含指定行数和列数的表格。insertLocation 的可取值为 'Before' 或 'After'。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > intersectWith(range: Range)|返回新范围作为此范围与另一范围的交集。此范围不变。|1.3|
|[range](/javascript/api/word/word.range)|_方法_ > split(delimiters: string, multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|使用分隔符将范围拆分为各个子范围。|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_属性_ > items|一组 range 对象。只读。|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_方法_ > getFirst()|获取此集合中的第一个范围。|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_方法_ > getItem(index: number)|按 range 对象在集合中的索引获取此对象。|1.3|
|[RequestContext](/javascript/api/word/word.requestcontext)|_方法_ > load(object: object, option: object)|使用参数指定的属性和选项填充在 JavaScript 层中创建的代理对象。 |1.3|
|[RequestContext](/javascript/api/word/word.requestcontext)|_方法_ > sync()|将请求队列提交到 Word 并返回一个 promise 对象，此对象可用于将其他操作链接起来。|1.3|
|[section](/javascript/api/word/word.section)|_方法_ > getNext()|获取下一个节。|1.3|
|[sectionCollection](/javascript/api/word/word.sectioncollection)|_方法_ > getFirst()|获取此集合中的第一个节。|1.3|
|[table](/javascript/api/word/word.table)|属性 > headerRowCount_ _|获取并设置标题行数。|1.3|
|[table](/javascript/api/word/word.table)|属性 > height_ _|获取表格的高度（以磅为单位）。只读。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > isUniform|指明所有表行是否一致。只读。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > nestingLevel|获取表格的嵌套级别。顶级表格的级别为 1。只读。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > rowCount|获取表格的行数。只读。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > shadingColor|获取并设置底纹色。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > style|获取或设置表格的样式名称。请对自定义样式和本地化样式名称使用此属性。若要使用可以在区域设置之间移植的嵌入样式，请参阅 "styleBuiltIn" 属性。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > styleBandedColumns|获取并设置表格是否有镶边列。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > styleBandedRows|获取并设置表格是否有镶边行。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > styleFirstColumn|获取并设置表格的第一列是否采用特殊样式。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > styleLastColumn|获取并设置表格的最后一列是否采用特殊样式。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > styleTotalRow|获取并设置表格是否具有采用特殊样式的总计行（最后一行）。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > values|以 2D Javascript 数组形式获取并设置表格中的文本值。|1.3|
|[table](/javascript/api/word/word.table)|_属性_ > width|获取并设置表格的宽度（以磅为单位）。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > font|获取字体。使用此关系可获取并设置字体名称、大小、颜色和其他属性。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > horizontalAlignment|获取并设置表格中每个单元格的水平对齐方式。可取值为 'left'、'centered'、'right'或 'justified'。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > paragraphAfter|获取表格之后的段落。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > paragraphBefore|获取表格之前的段落。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > parentBody|获取表格的父正文。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > parentContentControl|获取包含表格的内容控件。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > parentTable|获取包含此表格的表格。如果未包含在表格中，则返回空对象。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > parentTableCell|获取包含此表格的表格。如果未包含在表格单元格中，则返回空对象。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > rows|获取所有表格行。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > styleBuiltIn|获取或设置表格的嵌入样式名称。请对可以在区域设置之间移植的嵌入样式使用此属性。若要使用自定义样式或本地化样式名称，请参阅 "style" 属性。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > tables|获取嵌套一级的子表格。只读。|1.3|
|[table](/javascript/api/word/word.table)|_关系_ > verticalAlignment|获取并设置表格中每个单元格的垂直对齐方式。可取值为 'top'、'center' 或 'bottom'。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > addColumns(insertLocation: InsertLocation, columnCount: number, values: string)|使用第一个或最后一个现有列作为模板，将列添加到表格的开头或结尾。此方法适用于一致的表格。字符串值（若指定）是在新插入的行中进行设置。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > addRows(insertLocation: InsertLocation, rowCount: number, values: string)|使用第一个或最后一个现有行作为模板，将行添加到表格的开头或结尾。字符串值（若指定）是在新插入的行中进行设置。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > autoFitContents()|自动调整表格列，以适应内容的宽度。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > autoFitWindow()|自动调整表列，以适应窗口的宽度。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > clear()|清除表格的内容。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > delete()|删除整个表格。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > deleteColumns(columnIndex: number, columnCount: number)|删除特定的列。此方法适用于均一的表格。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > deleteRows(rowIndex: number, rowCount: number)|删除特定的行。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > distributeColumns()|将列设置为等宽。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > distributeRows()|将行设置为等高。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > getBorder(borderLocation: BorderLocation)|获取指定边框的边框样式。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > getCell(rowIndex: number, cellIndex: number)|获取指定行和列的表格单元格。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|获取单元格填充（以磅为单位）。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > getNext()|获取下一个表格。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > getRange(rangeLocation: RangeLocation)|获取包含此表格的范围，或此表格开头或结尾的范围。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > insertContentControl()|在表格中插入内容控件。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > insertParagraph(paragraphText: string, insertLocation: InsertLocation)|在指定位置插入段落。insertLocation 值可以为 'Before' 或 'After'。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|插入包含指定行数和列数的表格。insertLocation 的可取值为 'Before' 或 'After'。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)|使用指定的 searchOptions 在 table 对象范围内执行搜索。搜索结果是一组 range 对象。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > select(selectionMode: SelectionMode)|选择表格或其开头或结尾位置，然后将 Word UI 导航到相应位置。|1.3|
|[table](/javascript/api/word/word.table)|_方法_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|设置单元格填充（以磅为单位）。|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_属性_ > color|以十六进制值或名称的形式获取或设置表边框颜色。|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_属性_ > width|获取或设置表边框的宽度（以磅为单位）。不适用于宽度固定的表边框类型。|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_关系_ >type|获取或设置表边框的类型。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_属性_ > cellIndex|获取单元格在行中的索引。只读。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_属性_ > columnWidth|获取并设置单元格的列宽度（以磅为单位）。此方法适用于均一的表格。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_属性_ > rowIndex|获取单元格所在行在表中的索引。只读。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_属性_ > shadingColor|获取或设置单元格的底纹色。按 "#RRGGBB" 格式或使用颜色名称指定颜色。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_属性_ > value|获取并设置单元格的文本。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_属性_ > width|获取单元格的宽度（以磅为单位）。只读。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_关系_ > body|获取单元格的 body 对象。只读。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_关系_ > horizontalAlignment|获取并设置单元格的水平对齐方式。可取值为 'left'、'centered'、'right' 或 'justified'。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_关系_ > parentRow|获取单元格的父行。只读。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_关系_ > parentTable|获取单元格的父表格。只读。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_关系_ > verticalAlignment|获取并设置单元格的垂直对齐方式。可取值为 'top'、'center' 或' bottom'。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > deleteColumn()|删除包含该单元格的列。此方法适用于均一的表格。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > deleteRow()|删除包含该单元格的行。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > getBorder(borderLocation: BorderLocation)|获取指定边框的边框样式。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|获取单元格填充（以磅为单位）。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > getNext()|获取下一个单元格。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > addColumns(insertLocation: InsertLocation, columnCount: number, values: string)|使用单元格的列作为模板，在单元格的左侧或右侧添加列。此方法适用于均一的表格。字符串值（若指定）是在新插入的行中进行设置。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > insertRows(insertLocation: InsertLocation, rowCount: number, values: string)|使用单元格的行作为模板，在单元格的上方或下方插入行。字符串值（若指定）是在新插入的行中进行设置。|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_方法_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|设置单元格填充（以磅为单位）。|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_属性_ > items|一组 tableCell 对象。只读。|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_方法_ > getFirst()|获取此集合中的第一个 tableCell。|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_方法_ > getItem(index: number)|按 tableCell 对象在集合中的索引获取此对象。|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_属性_ > items|table 对象的集合。只读。|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_方法_ > getFirst()|获取此集合中的第一个表格。|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_方法_ > getItem(index: number)|按 table 对象在集合中的索引获取此对象。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_属性_ > cellCount|获取行单元格数。只读。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_属性_ > isHeader|检查该行是否为标题行。只读。若要设置标题行数，请对 Table 对象使用 HeaderRowCount。只读。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_属性_ > preferredHeight|获取并设置行的首选高度（以磅为单位）。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_属性_ > rowIndex|获取行在其父表中的索引。只读。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_属性_ > shadingColor|获取并设置底纹色。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_属性_ > values|以 1D Javascript 数组的形式获取并设置行中的文本值。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_关系_ > cells|获取单元格。只读。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_关系_ > font|获取字体。使用此关系可获取并设置字体名称、大小、颜色和其他属性。只读。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_关系_ > horizontalAlignment|获取并设置行中每个单元格的水平对齐方式。可取值为 'left'、'centered'、'right' 或 'justified'。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_关系_ > parentTable|获取父表。只读。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_关系_ > verticalAlignment|获取并设置行中单元格的垂直对齐方式。可取值为 'top'、'center' 或 'bottom'。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > clear()|清除行的内容。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > delete()|删除整行。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > getBorder(borderLocation: BorderLocation)|获取行中单元格的边框样式。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|获取单元格填充（以磅为单位）。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > getNext()|获取下一行。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > insertRows(insertLocation: InsertLocation, rowCount: number, values: string)|使用此行作为模板插入行。如果已指定值，将该值插入新行。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)|使用指定的 searchOptions 在行范围内执行搜索。搜索结果是一组 range 对象。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > select(selectionMode: SelectionMode)|选择行，然后将 Word UI 导航到相应位置。|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_方法_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|设置单元格填充（以磅为单位）。|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_属性_ > items|tableRow 对象的集合。只读。|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_方法_ > getFirst()|获取此集合中的第一行。|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_方法_ > getItem(index: number)|按 tableRow 对象在集合中的索引获取此对象。|1.3|


## <a name="whats-new-in-word-javascript-api-12"></a>Word JavaScript API 1.2 的最近更新

下面介绍了要求集 1.2 中 Word JavaScript API 的新增内容。 

|对象| 最近更新| 说明|要求集|
|:-----|-----|:----|:----|
|[contentControl](/javascript/api/word/word.contentcontrol)|_方法_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)|将内联图像插入到内容控件中的指定位置。insertLocation 值可以为 'Replace'、'Start' 或 'End'。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_关系_ > paragraph|获取包含内联图像的父段落。只读。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > delete()|从文档中删除内联图像。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > insertBreak (breakType: BreakType，insertLocation: InsertLocation)|在主文档的指定位置插入分隔符。insertLocation 的可取值为 'Before' 或 'After'。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > insertFileFromBase64(base64File: string, insertLocation: InsertLocation)|在指定位置插入文档。insertLocation 的可取值为 'Before' 或 'After'。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > insertHtml(html: string, insertLocation: InsertLocation)|在指定位置插入 HTML。insertLocation 值可以为 'Before' 或 'After'。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation|在指定位置插入 inlinePicture。insertLocation 的可取值为 'Replace'、'Before' 或 'After'。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > insertOoxml(ooxml: string, insertLocation: InsertLocation)|在指定位置插入 OOXML。insertLocation 值可以为 ’Before‘ 或 ’After‘。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > insertParagraph(paragraphText: string, insertLocation: InsertLocation)|在指定位置插入段落。insertLocation 值可以为 'Before' 或 'After'。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > insertText(text: string, insertLocation: InsertLocation)|在指定位置插入文本。insertLocation 的可取值为 ’Before‘ 或 ’After‘。|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_方法_ > select(selectionMode: SelectionMode)|选择内联图像。这会导致 Word 滚动到选定内容。|1.2|
|[range](/javascript/api/word/word.range)|_关系_ > inlinePictures|获取范围中的一组 inlinePicture 对象。只读。|1.2|
|[range](/javascript/api/word/word.range)|_方法_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)|在指定位置插入图片。insertLocation 的可取值为 'Replace'、'Start'、'End'、'Before' 或 'After'。|1.2|

## <a name="word-javascript-api-11"></a>Word JavaScript API 1.1

Word JavaScript API 1.1 是第一版 API。有关 API 的详细信息，请参阅 [Word JavaScript API](/javascript/api/word) 参考主题。 

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 宿主和 API 要求](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 加载项 XML 清单](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)

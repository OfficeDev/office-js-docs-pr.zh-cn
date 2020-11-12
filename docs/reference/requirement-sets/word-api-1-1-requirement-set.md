---
title: Word JavaScript API 要求集1。1
description: 有关 WordApi 1.1 要求集的详细信息
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 371638c18cff882f2b3907f1adedb6748761cc0c
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996436"
---
# <a name="whats-new-in-word-javascript-api-11"></a>Word JavaScript API 1.1 中的新增功能

WordApi 1.1 是 Word JavaScript API 的第一个要求集。 它是 Word 2016 仅支持的 Word API 要求集。

## <a name="api-list"></a>API 列表

下表列出了 Word JavaScript API 要求集1.1 中的 Api。 若要查看 Word JavaScript API 要求集1.1 支持的所有 Api 的 API 参考文档，请参阅 [要求集1.1 中的 Word api](/javascript/api/word?view=word-js-1.1&preserve-view=true)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[正文](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear--)|清除 body 对象的内容。|
||[getHtml()](/javascript/api/word/word.body#gethtml--)|获取 body 对象的 HTML 表示形式。|
||[getOoxml()](/javascript/api/word/word.body#getooxml--)|获取 body 对象的 OOXML (Office Open XML) 表示形式。|
||[ignorePunct](/javascript/api/word/word.body#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.body#ignorespace)||
||[insertBreak (breakType： BreakType，insertLocation： Word. InsertLocation) ](/javascript/api/word/word.body#insertbreak-breaktype--insertlocation-)|在主文档的指定位置插入分隔符。|
||[insertContentControl()](/javascript/api/word/word.body#insertcontentcontrol--)|使用富文本内容控件封装 body 对象。|
||[insertFileFromBase64 (base64File： string，insertLocation： InsertLocation) ](/javascript/api/word/word.body#insertfilefrombase64-base64file--insertlocation-)|将文档插入到正文中的指定位置。|
||[insertHtml (html： string，insertLocation： InsertLocation) ](/javascript/api/word/word.body#inserthtml-html--insertlocation-)|在指定位置插入 HTML。|
||[insertOoxml (ooxml： string，insertLocation： InsertLocation) ](/javascript/api/word/word.body#insertooxml-ooxml--insertlocation-)|在指定位置插入 OOXML。|
||[insertParagraph (paragraphText： string，insertLocation： InsertLocation) ](/javascript/api/word/word.body#insertparagraph-paragraphtext--insertlocation-)|在指定位置插入段落。|
||[insertText (text： string，insertLocation： InsertLocation) ](/javascript/api/word/word.body#inserttext-text--insertlocation-)|将文本插入到正文中的指定位置。|
||[matchCase](/javascript/api/word/word.body#matchcase)||
||[matchPrefix](/javascript/api/word/word.body#matchprefix)||
||[matchSuffix](/javascript/api/word/word.body#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.body#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.body#matchwildcards)||
||[contentControls](/javascript/api/word/word.body#contentcontrols)|获取正文中的格式文本内容控件对象的集合。|
||[font](/javascript/api/word/word.body#font)|获取正文的文本格式。|
||[inlinePictures](/javascript/api/word/word.body#inlinepictures)|获取正文中 InlinePicture 对象的集合。|
||[paragraphs](/javascript/api/word/word.body#paragraphs)|获取正文中段落对象的集合。|
||[parentContentControl](/javascript/api/word/word.body#parentcontentcontrol)|获取包含正文的内容控件。|
||[text](/javascript/api/word/word.body#text)|获取正文的文本。|
||[search (searchText： string，searchOptions？： Word. SearchOptions \| {ignorePunct？： Boolean ignoreSpace？： Boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： Boolean matchWholeWord？： Boolean matchWildcards？： boolean} ) ](/javascript/api/word/word.body#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|在 body 对象的作用域上使用指定的 SearchOptions 执行搜索。|
||[选择 (selectionMode？： SelectionMode) ](/javascript/api/word/word.body#select-selectionmode-)|选择正文并在 Word UI 中进行浏览。|
||[style](/javascript/api/word/word.body#style)|获取或设置正文的样式名称。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[自然](/javascript/api/word/word.contentcontrol#appearance)|获取或设置内容控件的外观。|
||[cannotDelete](/javascript/api/word/word.contentcontrol#cannotdelete)|获取或设置指示用户是否可以删除内容控件的值。|
||[cannotEdit](/javascript/api/word/word.contentcontrol#cannotedit)|获取或设置指示用户是否可以编辑内容控件的内容的值。|
||[clear()](/javascript/api/word/word.contentcontrol#clear--)|清除内容控件的内容。|
||[color](/javascript/api/word/word.contentcontrol#color)|获取或设置内容控件的颜色。|
||[删除 (keepContent： boolean) ](/javascript/api/word/word.contentcontrol#delete-keepcontent-)|删除内容控件及其内容。|
||[getHtml()](/javascript/api/word/word.contentcontrol#gethtml--)|获取内容控件对象的 HTML 表示形式。|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getooxml--)|获取内容控件对象的 Office Open XML (OOXML) 表示形式。|
||[ignorePunct](/javascript/api/word/word.contentcontrol#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#ignorespace)||
||[insertBreak (breakType： BreakType，insertLocation： Word. InsertLocation) ](/javascript/api/word/word.contentcontrol#insertbreak-breaktype--insertlocation-)|在主文档的指定位置插入分隔符。|
||[insertFileFromBase64 (base64File： string，insertLocation： InsertLocation) ](/javascript/api/word/word.contentcontrol#insertfilefrombase64-base64file--insertlocation-)|将文档插入到内容控件中的指定位置。|
||[insertHtml (html： string，insertLocation： InsertLocation) ](/javascript/api/word/word.contentcontrol#inserthtml-html--insertlocation-)|将 HTML 插入到内容控件中的指定位置。|
||[insertOoxml (ooxml： string，insertLocation： InsertLocation) ](/javascript/api/word/word.contentcontrol#insertooxml-ooxml--insertlocation-)|将 OOXML 插入到内容控件中的指定位置。|
||[insertParagraph (paragraphText： string，insertLocation： InsertLocation) ](/javascript/api/word/word.contentcontrol#insertparagraph-paragraphtext--insertlocation-)|在指定位置插入段落。|
||[insertText (text： string，insertLocation： InsertLocation) ](/javascript/api/word/word.contentcontrol#inserttext-text--insertlocation-)|将文本插入到内容控件中的指定位置。|
||[matchCase](/javascript/api/word/word.contentcontrol#matchcase)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#matchprefix)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#matchwildcards)||
||[placeholderText](/javascript/api/word/word.contentcontrol#placeholdertext)|获取或设置内容控件的占位符文本。|
||[contentControls](/javascript/api/word/word.contentcontrol#contentcontrols)|获取内容控件中的内容控件对象的集合。|
||[font](/javascript/api/word/word.contentcontrol#font)|获取内容控件的文本格式。|
||[id](/javascript/api/word/word.contentcontrol#id)|获取表示内容控件标识符的整数。|
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinepictures)|获取内容控件中的 inlinePicture 对象的集合。|
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|获取内容控件中的 paragraph 对象的集合。|
||[parentContentControl](/javascript/api/word/word.contentcontrol#parentcontentcontrol)|获取包含此内容控件的内容控件。|
||[text](/javascript/api/word/word.contentcontrol#text)|获取内容控件的文本。|
||[type](/javascript/api/word/word.contentcontrol#type)|获取内容控件的类型。|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#removewhenedited)|获取或设置指示内容控件在编辑后是否可以删除的值。|
||[search (searchText： string，searchOptions？： Word. SearchOptions \| {ignorePunct？： Boolean ignoreSpace？： Boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： Boolean matchWholeWord？： Boolean matchWildcards？： boolean} ) ](/javascript/api/word/word.contentcontrol#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|在内容控件对象的范围内使用指定的 SearchOptions 执行搜索。|
||[选择 (selectionMode？： SelectionMode) ](/javascript/api/word/word.contentcontrol#select-selectionmode-)|选择内容控件。|
||[style](/javascript/api/word/word.contentcontrol#style)|获取或设置内容控件的样式名称。|
||[tag](/javascript/api/word/word.contentcontrol#tag)|获取或设置用于标识内容控件的标记。|
||[title](/javascript/api/word/word.contentcontrol#title)|获取或设置内容控件的标题。|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyid-id-)|按其标识符获取内容控件。|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getbytag-tag-)|获取具有指定标记的内容控件。|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getbytitle-title-)|获取具有指定标题的内容控件。|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#getitem-index-)|按其在集合中的索引获取内容控件。|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|获取此集合中已加载的子项。|
|[Document](/javascript/api/word/word.document)|[getSelection ( # B1 ](/javascript/api/word/word.document#getselection--)|获取文档的当前选定内容。|
||[body](/javascript/api/word/word.document#body)|获取文档的正文对象。|
||[contentControls](/javascript/api/word/word.document#contentcontrols)|获取文档中的内容控件对象的集合。|
||[保存](/javascript/api/word/word.document#saved)|指示是否已保存在文档中所做的更改。|
||[sections](/javascript/api/word/word.document#sections)|获取文档中的节对象的集合。|
||[save()](/javascript/api/word/word.document#save--)|保存文档。|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|获取或设置表示字体是否为粗体的值。|
||[color](/javascript/api/word/word.font#color)|获取或设置指定字体的颜色。|
||[doubleStrikeThrough](/javascript/api/word/word.font#doublestrikethrough)|获取或设置一个值，该值指示字体是否具有双删除线。|
||[highlightColor](/javascript/api/word/word.font#highlightcolor)|获取或设置突出显示颜色。|
||[italic](/javascript/api/word/word.font#italic)|获取或设置表示字体是否为斜体的值。|
||[name](/javascript/api/word/word.font#name)|获取或设置表示字体名称的值。|
||[size](/javascript/api/word/word.font#size)|获取或设置表示字体大小（以磅值表示）的值。|
||[删除](/javascript/api/word/word.font#strikethrough)|获取或设置一个值，该值指示字体是否具有删除线。|
||[subscript](/javascript/api/word/word.font#subscript)|获取或设置表示字体是否为下标的值。|
||[superscript](/javascript/api/word/word.font#superscript)|获取或设置表示字体是否为上标的值。|
||[underline](/javascript/api/word/word.font#underline)|获取或设置表示字体的下划线类型的值。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#alttextdescription)|获取或设置一个字符串，表示与嵌入式图像相关联的可选文字。|
||[altTextTitle](/javascript/api/word/word.inlinepicture#alttexttitle)|获取或设置包含嵌入式图像的标题的字符串。|
||[getBase64ImageSrc ( # B1 ](/javascript/api/word/word.inlinepicture#getbase64imagesrc--)|获取嵌入式图像的 base64 编码的字符串表示。|
||[height](/javascript/api/word/word.inlinepicture#height)|获取或设置描述嵌入式图像的高度的数字。|
||[hyperlink](/javascript/api/word/word.inlinepicture#hyperlink)|获取或设置图像上的超链接。|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertcontentcontrol--)|使用富文本内容控件封装嵌入式图像。|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockaspectratio)|获取或设置指示在您调整嵌入式图像大小时其是否保留原始比例的值。|
||[parentContentControl](/javascript/api/word/word.inlinepicture#parentcontentcontrol)|获取包含嵌入式图像的内容控件。|
||[width](/javascript/api/word/word.inlinepicture#width)|获取或设置描述嵌入式图像的宽度的数字。|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|获取此集合中已加载的子项。|
|[Paragraph](/javascript/api/word/word.paragraph)|[对齐方式](/javascript/api/word/word.paragraph#alignment)|获取或设置段落的对齐方式。|
||[clear()](/javascript/api/word/word.paragraph#clear--)|清除 paragraph 对象的内容。|
||[delete()](/javascript/api/word/word.paragraph#delete--)|从文档中删除段落及其内容。|
||[firstLineIndent](/javascript/api/word/word.paragraph#firstlineindent)|获取或设置首行缩进或悬挂缩进的大小（以磅值表示）。|
||[getHtml()](/javascript/api/word/word.paragraph#gethtml--)|获取段落对象的 HTML 表示形式。|
||[getOoxml()](/javascript/api/word/word.paragraph#getooxml--)|获取 paragraph 对象的 Office Open XML (OOXML) 表示形式。|
||[ignorePunct](/javascript/api/word/word.paragraph#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.paragraph#ignorespace)||
||[insertBreak (breakType： BreakType，insertLocation： Word. InsertLocation) ](/javascript/api/word/word.paragraph#insertbreak-breaktype--insertlocation-)|在主文档的指定位置插入分隔符。|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertcontentcontrol--)|使用富文本内容控件封装 paragraph 对象。|
||[insertFileFromBase64 (base64File： string，insertLocation： InsertLocation) ](/javascript/api/word/word.paragraph#insertfilefrombase64-base64file--insertlocation-)|将文档插入到指定位置的段落中。|
||[insertHtml (html： string，insertLocation： InsertLocation) ](/javascript/api/word/word.paragraph#inserthtml-html--insertlocation-)|将 HTML 插入到段落中的指定位置。|
||[insertInlinePictureFromBase64 (base64EncodedImage： string，insertLocation： InsertLocation) ](/javascript/api/word/word.paragraph#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|将图片插入到段落中的指定位置。|
||[insertOoxml (ooxml： string，insertLocation： InsertLocation) ](/javascript/api/word/word.paragraph#insertooxml-ooxml--insertlocation-)|将 OOXML 插入到段落中的指定位置。|
||[insertParagraph (paragraphText： string，insertLocation： InsertLocation) ](/javascript/api/word/word.paragraph#insertparagraph-paragraphtext--insertlocation-)|在指定位置插入段落。|
||[insertText (text： string，insertLocation： InsertLocation) ](/javascript/api/word/word.paragraph#inserttext-text--insertlocation-)|将文本插入到段落中的指定位置。|
||[leftIndent](/javascript/api/word/word.paragraph#leftindent)|获取或设置段落的向左缩进值（以磅值表示）。|
||[lineSpacing](/javascript/api/word/word.paragraph#linespacing)|获取或设置指定段落的行间距（以磅值表示）。|
||[lineUnitAfter](/javascript/api/word/word.paragraph#lineunitafter)|获取或设置段落后面的网格线中的间距量。|
||[lineUnitBefore](/javascript/api/word/word.paragraph#lineunitbefore)|获取或设置段落前面的网格线中的间隔量。|
||[matchCase](/javascript/api/word/word.paragraph#matchcase)||
||[matchPrefix](/javascript/api/word/word.paragraph#matchprefix)||
||[matchSuffix](/javascript/api/word/word.paragraph#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.paragraph#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.paragraph#matchwildcards)||
||[outlineLevel](/javascript/api/word/word.paragraph#outlinelevel)|获取或设置段落的大纲级别。|
||[contentControls](/javascript/api/word/word.paragraph#contentcontrols)|获取段落中的内容控件对象的集合。|
||[font](/javascript/api/word/word.paragraph#font)|获取段落的文本格式。|
||[inlinePictures](/javascript/api/word/word.paragraph#inlinepictures)|获取段落中 InlinePicture 对象的集合。|
||[parentContentControl](/javascript/api/word/word.paragraph#parentcontentcontrol)|获取包含段落的内容控件。|
||[text](/javascript/api/word/word.paragraph#text)|获取段落的文本。|
||[rightIndent](/javascript/api/word/word.paragraph#rightindent)|获取或设置段落的向右缩进值（以磅值表示）。|
||[search (searchText： string，searchOptions？： Word. SearchOptions \| {ignorePunct？： Boolean ignoreSpace？： Boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： Boolean matchWholeWord？： Boolean matchWildcards？： boolean} ) ](/javascript/api/word/word.paragraph#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|在段落对象的作用域上使用指定的 SearchOptions 执行搜索。|
||[选择 (selectionMode？： SelectionMode) ](/javascript/api/word/word.paragraph#select-selectionmode-)|选择并在 Word UI 中导航到段落。|
||[spaceAfter](/javascript/api/word/word.paragraph#spaceafter)|获取或设置段落后面的间距（以磅值表示）。|
||[spaceBefore](/javascript/api/word/word.paragraph#spacebefore)|获取或设置段落前面的间距（以磅值表示）。|
||[style](/javascript/api/word/word.paragraph#style)|获取或设置段落的样式名称。|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|获取此集合中已加载的子项。|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear--)|清除 range 对象的内容。|
||[delete()](/javascript/api/word/word.range#delete--)|从文档中删除区域及其内容。|
||[getHtml()](/javascript/api/word/word.range#gethtml--)|获取 range 对象的 HTML 表示形式。|
||[getOoxml()](/javascript/api/word/word.range#getooxml--)|获取 range 对象的 OOXML 表示形式。|
||[ignorePunct](/javascript/api/word/word.range#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.range#ignorespace)||
||[insertBreak (breakType： BreakType，insertLocation： Word. InsertLocation) ](/javascript/api/word/word.range#insertbreak-breaktype--insertlocation-)|在主文档的指定位置插入分隔符。|
||[insertContentControl()](/javascript/api/word/word.range#insertcontentcontrol--)|使用富文本内容控件封装 range 对象。|
||[insertFileFromBase64 (base64File： string，insertLocation： InsertLocation) ](/javascript/api/word/word.range#insertfilefrombase64-base64file--insertlocation-)|在指定位置插入 document。|
||[insertHtml (html： string，insertLocation： InsertLocation) ](/javascript/api/word/word.range#inserthtml-html--insertlocation-)|在指定位置插入 HTML。|
||[insertOoxml (ooxml： string，insertLocation： InsertLocation) ](/javascript/api/word/word.range#insertooxml-ooxml--insertlocation-)|在指定位置插入 OOXML。|
||[insertParagraph (paragraphText： string，insertLocation： InsertLocation) ](/javascript/api/word/word.range#insertparagraph-paragraphtext--insertlocation-)|在指定位置插入段落。|
||[insertText (text： string，insertLocation： InsertLocation) ](/javascript/api/word/word.range#inserttext-text--insertlocation-)|在指定位置插入文本。|
||[matchCase](/javascript/api/word/word.range#matchcase)||
||[matchPrefix](/javascript/api/word/word.range#matchprefix)||
||[matchSuffix](/javascript/api/word/word.range#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.range#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.range#matchwildcards)||
||[contentControls](/javascript/api/word/word.range#contentcontrols)|获取范围中的内容控件对象的集合。|
||[font](/javascript/api/word/word.range#font)|获取区域的文本格式。|
||[paragraphs](/javascript/api/word/word.range#paragraphs)|获取范围中的段落对象的集合。|
||[parentContentControl](/javascript/api/word/word.range#parentcontentcontrol)|获取包含该范围的内容控件。|
||[text](/javascript/api/word/word.range#text)|获取区域的文本。|
||[search (searchText： string，searchOptions？： Word. SearchOptions \| {ignorePunct？： Boolean ignoreSpace？： Boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： Boolean matchWholeWord？： Boolean matchWildcards？： boolean} ) ](/javascript/api/word/word.range#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|在 range 对象的作用域上使用指定的 SearchOptions 执行搜索。|
||[选择 (selectionMode？： SelectionMode) ](/javascript/api/word/word.range#select-selectionmode-)|选择并在 Word UI 中导航到区域。|
||[style](/javascript/api/word/word.range#style)|获取或设置区域的样式名称。|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|获取此集合中已加载的子项。|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#ignorepunct)|获取或设置指示是否忽略单词之间的所有标点符号的值。|
||[ignoreSpace](/javascript/api/word/word.searchoptions#ignorespace)|获取或设置一个值，该值指示是否忽略单词之间的所有空格。|
||[matchCase](/javascript/api/word/word.searchoptions#matchcase)|获取或设置指示是否执行区分大小写的搜索的值。|
||[matchPrefix](/javascript/api/word/word.searchoptions#matchprefix)|获取或设置指示是否匹配以搜索字符串开头的单词。|
||[matchSuffix](/javascript/api/word/word.searchoptions#matchsuffix)|获取或设置指示是否匹配以搜索字符串结尾的单词。|
||[matchWholeWord](/javascript/api/word/word.searchoptions#matchwholeword)|获取或设置指示是否只查找整个单词，而不查找长单词的一部分的值。|
||[matchWildcards](/javascript/api/word/word.searchoptions#matchwildcards)|获取或设置指示搜索是否使用特殊搜索操作符执行的值。|
|[Section](/javascript/api/word/word.section)|[getFooter (类型： HeaderFooterType) ](/javascript/api/word/word.section#getfooter-type-)|获取节的页脚之一。|
||[getHeader (类型： HeaderFooterType) ](/javascript/api/word/word.section#getheader-type-)|获取节的标头之一。|
||[body](/javascript/api/word/word.section#body)|获取节的 body 对象。|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Word JavaScript API 要求集](word-api-requirement-sets.md)

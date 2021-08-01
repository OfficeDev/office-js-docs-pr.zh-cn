---
title: Word JavaScript API 要求集 1.1
description: 有关 WordApi 1.1 要求集的详细信息
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 80914cd0804600e7987408ce3a3de8a94e6fec29
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671595"
---
# <a name="whats-new-in-word-javascript-api-11"></a>Word JavaScript API 1.1 的新增功能

WordApi 1.1 是 Word JavaScript API 的第一个要求集。 这是唯一受 Word API 要求集支持Word 2016。

## <a name="api-list"></a>API 列表

下表列出了 Word JavaScript API 要求集 1.1 中的 API。 若要查看受 Word JavaScript API 要求集 1.1 支持的所有 API 的 API 参考文档，请参阅要求集 [1.1](/javascript/api/word?view=word-js-1.1&preserve-view=true)中的 Word API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[正文](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear__)|清除 body 对象的内容。|
||[getHtml()](/javascript/api/word/word.body#getHtml__)|获取 body 对象的 HTML 表示形式。|
||[getOoxml()](/javascript/api/word/word.body#getOoxml__)|获取 body 对象的 OOXML (Office Open XML) 表示形式。|
||[ignorePunct](/javascript/api/word/word.body#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.body#ignoreSpace)||
||[insertBreak (breakType： Word.BreakType， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.body#insertBreak_breakType__insertLocation_)|在主文档的指定位置插入分隔符。|
||[insertContentControl()](/javascript/api/word/word.body#insertContentControl__)|使用富文本内容控件封装 body 对象。|
||[insertFileFromBase64 (base64File： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.body#insertFileFromBase64_base64File__insertLocation_)|将文档插入到正文中的指定位置。|
||[insertHtml (html： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.body#insertHtml_html__insertLocation_)|在指定位置插入 HTML。|
||[insertOoxml (ooxml： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.body#insertOoxml_ooxml__insertLocation_)|在指定位置插入 OOXML。|
||[insertParagraph (paragraphText： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.body#insertParagraph_paragraphText__insertLocation_)|在指定位置插入段落。|
||[insertText (text： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.body#insertText_text__insertLocation_)|将文本插入到正文中的指定位置。|
||[matchCase](/javascript/api/word/word.body#matchCase)||
||[matchPrefix](/javascript/api/word/word.body#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.body#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.body#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.body#matchWildcards)||
||[contentControls](/javascript/api/word/word.body#contentControls)|获取正文中的格式文本内容控件对象的集合。|
||[font](/javascript/api/word/word.body#font)|获取正文的文本格式。|
||[inlinePictures](/javascript/api/word/word.body#inlinePictures)|获取正文中的 InlinePicture 对象的集合。|
||[paragraphs](/javascript/api/word/word.body#paragraphs)|获取 body 中的 paragraph 对象的集合。|
||[parentContentControl](/javascript/api/word/word.body#parentContentControl)|获取包含正文的内容控件。|
||[text](/javascript/api/word/word.body#text)|获取正文的文本。|
||[search (searchText： string， searchOptions？： Word.SearchOptions \| { ignorePunct？： boolean ignoreSpace？： boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： boolean matchWholeWord？： boolean matchWildcards？： boolean }) ](/javascript/api/word/word.body#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|在 body 对象范围内使用指定的 SearchOptions 执行搜索。|
||[select (selectionMode？： Word.SelectionMode) ](/javascript/api/word/word.body#select_selectionMode_)|选择正文并在 Word UI 中进行浏览。|
||[style](/javascript/api/word/word.body#style)|获取或设置正文的样式名称。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[appearance](/javascript/api/word/word.contentcontrol#appearance)|获取或设置内容控件的外观。|
||[cannotDelete](/javascript/api/word/word.contentcontrol#cannotDelete)|获取或设置指示用户是否可以删除内容控件的值。|
||[cannotEdit](/javascript/api/word/word.contentcontrol#cannotEdit)|获取或设置指示用户是否可以编辑内容控件的内容的值。|
||[clear()](/javascript/api/word/word.contentcontrol#clear__)|清除内容控件的内容。|
||[color](/javascript/api/word/word.contentcontrol#color)|获取或设置内容控件的颜色。|
||[delete (keepContent： boolean) ](/javascript/api/word/word.contentcontrol#delete_keepContent_)|删除内容控件及其内容。|
||[getHtml()](/javascript/api/word/word.contentcontrol#getHtml__)|获取内容控件对象的 HTML 表示形式。|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getOoxml__)|获取内容控件对象的 Office Open XML (OOXML) 表示形式。|
||[ignorePunct](/javascript/api/word/word.contentcontrol#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#ignoreSpace)||
||[insertBreak (breakType： Word.BreakType， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.contentcontrol#insertBreak_breakType__insertLocation_)|在主文档的指定位置插入分隔符。|
||[insertFileFromBase64 (base64File： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.contentcontrol#insertFileFromBase64_base64File__insertLocation_)|将文档插入到内容控件中的指定位置。|
||[insertHtml (html： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.contentcontrol#insertHtml_html__insertLocation_)|将 HTML 插入到内容控件中的指定位置。|
||[insertOoxml (ooxml： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.contentcontrol#insertOoxml_ooxml__insertLocation_)|将 OOXML 插入到位于指定位置的内容控件中。|
||[insertParagraph (paragraphText： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.contentcontrol#insertParagraph_paragraphText__insertLocation_)|在指定位置插入段落。|
||[insertText (text： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.contentcontrol#insertText_text__insertLocation_)|将文本插入到内容控件中的指定位置。|
||[matchCase](/javascript/api/word/word.contentcontrol#matchCase)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#matchWildcards)||
||[placeholderText](/javascript/api/word/word.contentcontrol#placeholderText)|获取或设置内容控件的占位符文本。|
||[contentControls](/javascript/api/word/word.contentcontrol#contentControls)|获取内容控件中的内容控件对象的集合。|
||[font](/javascript/api/word/word.contentcontrol#font)|获取内容控件的文本格式。|
||[id](/javascript/api/word/word.contentcontrol#id)|获取表示内容控件标识符的整数。|
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinePictures)|获取内容控件中的 inlinePicture 对象的集合。|
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|获取内容控件中的 paragraph 对象的集合。|
||[parentContentControl](/javascript/api/word/word.contentcontrol#parentContentControl)|获取包含此内容控件的内容控件。|
||[text](/javascript/api/word/word.contentcontrol#text)|获取内容控件的文本。|
||[type](/javascript/api/word/word.contentcontrol#type)|获取内容控件的类型。|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#removeWhenEdited)|获取或设置指示内容控件在编辑后是否可以删除的值。|
||[search (searchText： string， searchOptions？： Word.SearchOptions \| { ignorePunct？： boolean ignoreSpace？： boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： boolean matchWholeWord？： boolean matchWildcards？： boolean }) ](/javascript/api/word/word.contentcontrol#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|在内容控件对象范围内使用指定的 SearchOptions 执行搜索。|
||[select (selectionMode？： Word.SelectionMode) ](/javascript/api/word/word.contentcontrol#select_selectionMode_)|选择内容控件。|
||[style](/javascript/api/word/word.contentcontrol#style)|获取或设置内容控件的样式名称。|
||[标记](/javascript/api/word/word.contentcontrol#tag)|获取或设置用于标识内容控件的标记。|
||[title](/javascript/api/word/word.contentcontrol#title)|获取或设置内容控件的标题。|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getById_id_)|按其标识符获取内容控件。|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getByTag_tag_)|获取具有指定标记的内容控件。|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getByTitle_title_)|获取具有指定标题的内容控件。|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#getItem_index_)|按内容控件在集合中的索引获取内容控件。|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|获取此集合中已加载的子项。|
|[Document](/javascript/api/word/word.document)|[getSelection () ](/javascript/api/word/word.document#getSelection__)|获取文档的当前选定内容。|
||[body](/javascript/api/word/word.document#body)|获取文档的 body 对象。|
||[contentControls](/javascript/api/word/word.document#contentControls)|获取文档中的内容控件对象的集合。|
||[saved](/javascript/api/word/word.document#saved)|指示是否已保存在文档中所做的更改。|
||[sections](/javascript/api/word/word.document#sections)|获取文档中 section 对象的集合。|
||[save()](/javascript/api/word/word.document#save__)|保存文档。|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|获取或设置表示字体是否为粗体的值。|
||[color](/javascript/api/word/word.font#color)|获取或设置指定字体的颜色。|
||[doubleStrikeThrough](/javascript/api/word/word.font#doubleStrikeThrough)|获取或设置一个值，该值指示字体是否具有双删除线。|
||[highlightColor](/javascript/api/word/word.font#highlightColor)|获取或设置突出显示颜色。|
||[italic](/javascript/api/word/word.font#italic)|获取或设置表示字体是否为斜体的值。|
||[名称](/javascript/api/word/word.font#name)|获取或设置表示字体名称的值。|
||[size](/javascript/api/word/word.font#size)|获取或设置表示字体大小（以磅值表示）的值。|
||[strikeThrough](/javascript/api/word/word.font#strikeThrough)|获取或设置一个值，该值指示字体是否有删除线。|
||[subscript](/javascript/api/word/word.font#subscript)|获取或设置表示字体是否为下标的值。|
||[superscript](/javascript/api/word/word.font#superscript)|获取或设置表示字体是否为上标的值。|
||[underline](/javascript/api/word/word.font#underline)|获取或设置表示字体的下划线类型的值。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#altTextDescription)|获取或设置一个字符串，该字符串代表与内联图像关联的可选文本。|
||[altTextTitle](/javascript/api/word/word.inlinepicture#altTextTitle)|获取或设置包含嵌入式图像的标题的字符串。|
||[getBase64ImageSrc () ](/javascript/api/word/word.inlinepicture#getBase64ImageSrc__)|获取嵌入式图像的 base64 编码的字符串表示。|
||[height](/javascript/api/word/word.inlinepicture#height)|获取或设置描述嵌入式图像的高度的数字。|
||[hyperlink](/javascript/api/word/word.inlinepicture#hyperlink)|获取或设置图像上的超链接。|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertContentControl__)|使用富文本内容控件封装嵌入式图像。|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockAspectRatio)|获取或设置指示在您调整嵌入式图像大小时其是否保留原始比例的值。|
||[parentContentControl](/javascript/api/word/word.inlinepicture#parentContentControl)|获取包含嵌入式图像的内容控件。|
||[width](/javascript/api/word/word.inlinepicture#width)|获取或设置描述嵌入式图像的宽度的数字。|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|获取此集合中已加载的子项。|
|[Paragraph](/javascript/api/word/word.paragraph)|[alignment](/javascript/api/word/word.paragraph#alignment)|获取或设置段落的对齐方式。|
||[clear()](/javascript/api/word/word.paragraph#clear__)|清除 paragraph 对象的内容。|
||[delete()](/javascript/api/word/word.paragraph#delete__)|从文档中删除段落及其内容。|
||[firstLineIndent](/javascript/api/word/word.paragraph#firstLineIndent)|获取或设置首行缩进或悬挂缩进的大小（以磅值表示）。|
||[getHtml()](/javascript/api/word/word.paragraph#getHtml__)|获取 paragraph 对象的 HTML 表示形式。|
||[getOoxml()](/javascript/api/word/word.paragraph#getOoxml__)|获取 paragraph 对象的 Office Open XML (OOXML) 表示形式。|
||[ignorePunct](/javascript/api/word/word.paragraph#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.paragraph#ignoreSpace)||
||[insertBreak (breakType： Word.BreakType， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.paragraph#insertBreak_breakType__insertLocation_)|在主文档的指定位置插入分隔符。|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertContentControl__)|使用富文本内容控件封装 paragraph 对象。|
||[insertFileFromBase64 (base64File： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.paragraph#insertFileFromBase64_base64File__insertLocation_)|将文档插入到段落中的指定位置。|
||[insertHtml (html： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.paragraph#insertHtml_html__insertLocation_)|将 HTML 插入到段落中的指定位置。|
||[insertInlinePictureFromBase64 (base64EncodedImage： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.paragraph#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|将图片插入到段落中的指定位置。|
||[insertOoxml (ooxml： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.paragraph#insertOoxml_ooxml__insertLocation_)|将 OOXML 插入到段落中的指定位置。|
||[insertParagraph (paragraphText： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.paragraph#insertParagraph_paragraphText__insertLocation_)|在指定位置插入段落。|
||[insertText (text： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.paragraph#insertText_text__insertLocation_)|将文本插入到段落中的指定位置。|
||[leftIndent](/javascript/api/word/word.paragraph#leftIndent)|获取或设置段落的向左缩进值（以磅值表示）。|
||[lineSpacing](/javascript/api/word/word.paragraph#lineSpacing)|获取或设置指定段落的行间距（以磅值表示）。|
||[lineUnitAfter](/javascript/api/word/word.paragraph#lineUnitAfter)|获取或设置段落后的间距，以网格线表示。|
||[lineUnitBefore](/javascript/api/word/word.paragraph#lineUnitBefore)|获取或设置段落前面的网格线中的间隔量。|
||[matchCase](/javascript/api/word/word.paragraph#matchCase)||
||[matchPrefix](/javascript/api/word/word.paragraph#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.paragraph#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.paragraph#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.paragraph#matchWildcards)||
||[outlineLevel](/javascript/api/word/word.paragraph#outlineLevel)|获取或设置段落的大纲级别。|
||[contentControls](/javascript/api/word/word.paragraph#contentControls)|获取段落中内容控件对象的集合。|
||[font](/javascript/api/word/word.paragraph#font)|获取段落的文本格式。|
||[inlinePictures](/javascript/api/word/word.paragraph#inlinePictures)|获取段落中的 InlinePicture 对象的集合。|
||[parentContentControl](/javascript/api/word/word.paragraph#parentContentControl)|获取包含段落的内容控件。|
||[text](/javascript/api/word/word.paragraph#text)|获取段落的文本。|
||[rightIndent](/javascript/api/word/word.paragraph#rightIndent)|获取或设置段落的向右缩进值（以磅值表示）。|
||[search (searchText： string， searchOptions？： Word.SearchOptions \| { ignorePunct？： boolean ignoreSpace？： boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： boolean matchWholeWord？： boolean matchWildcards？： boolean }) ](/javascript/api/word/word.paragraph#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|使用指定的 SearchOptions 搜索段落对象的范围。|
||[select (selectionMode？： Word.SelectionMode) ](/javascript/api/word/word.paragraph#select_selectionMode_)|选择并在 Word UI 中导航到段落。|
||[spaceAfter](/javascript/api/word/word.paragraph#spaceAfter)|获取或设置段落后面的间距（以磅值表示）。|
||[spaceBefore](/javascript/api/word/word.paragraph#spaceBefore)|获取或设置段落前面的间距（以磅值表示）。|
||[style](/javascript/api/word/word.paragraph#style)|获取或设置段落的样式名称。|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|获取此集合中已加载的子项。|
|[区域](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear__)|清除 range 对象的内容。|
||[delete()](/javascript/api/word/word.range#delete__)|从文档中删除区域及其内容。|
||[getHtml()](/javascript/api/word/word.range#getHtml__)|获取 range 对象的 HTML 表示形式。|
||[getOoxml()](/javascript/api/word/word.range#getOoxml__)|获取 range 对象的 OOXML 表示形式。|
||[ignorePunct](/javascript/api/word/word.range#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.range#ignoreSpace)||
||[insertBreak (breakType： Word.BreakType， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.range#insertBreak_breakType__insertLocation_)|在主文档的指定位置插入分隔符。|
||[insertContentControl()](/javascript/api/word/word.range#insertContentControl__)|使用富文本内容控件封装 range 对象。|
||[insertFileFromBase64 (base64File： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.range#insertFileFromBase64_base64File__insertLocation_)|在指定位置插入 document。|
||[insertHtml (html： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.range#insertHtml_html__insertLocation_)|在指定位置插入 HTML。|
||[insertOoxml (ooxml： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.range#insertOoxml_ooxml__insertLocation_)|在指定位置插入 OOXML。|
||[insertParagraph (paragraphText： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.range#insertParagraph_paragraphText__insertLocation_)|在指定位置插入段落。|
||[insertText (text： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.range#insertText_text__insertLocation_)|在指定位置插入文本。|
||[matchCase](/javascript/api/word/word.range#matchCase)||
||[matchPrefix](/javascript/api/word/word.range#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.range#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.range#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.range#matchWildcards)||
||[contentControls](/javascript/api/word/word.range#contentControls)|获取范围中的内容控件对象的集合。|
||[font](/javascript/api/word/word.range#font)|获取区域的文本格式。|
||[paragraphs](/javascript/api/word/word.range#paragraphs)|获取 range 中的 paragraph 对象的集合。|
||[parentContentControl](/javascript/api/word/word.range#parentContentControl)|获取包含该范围的内容控件。|
||[text](/javascript/api/word/word.range#text)|获取区域的文本。|
||[search (searchText： string， searchOptions？： Word.SearchOptions \| { ignorePunct？： boolean ignoreSpace？： boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： boolean matchWholeWord？： boolean matchWildcards？： boolean }) ](/javascript/api/word/word.range#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|使用指定的 SearchOptions 搜索 range 对象的范围。|
||[select (selectionMode？： Word.SelectionMode) ](/javascript/api/word/word.range#select_selectionMode_)|选择并在 Word UI 中导航到区域。|
||[style](/javascript/api/word/word.range#style)|获取或设置范围的样式名称。|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|获取此集合中已加载的子项。|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#ignorePunct)|获取或设置指示是否忽略单词之间的所有标点符号的值。|
||[ignoreSpace](/javascript/api/word/word.searchoptions#ignoreSpace)|获取或设置一个值，该值指示是否忽略单词之间的所有空格。|
||[matchCase](/javascript/api/word/word.searchoptions#matchCase)|获取或设置指示是否执行区分大小写的搜索的值。|
||[matchPrefix](/javascript/api/word/word.searchoptions#matchPrefix)|获取或设置指示是否匹配以搜索字符串开头的单词。|
||[matchSuffix](/javascript/api/word/word.searchoptions#matchSuffix)|获取或设置指示是否匹配以搜索字符串结尾的单词。|
||[matchWholeWord](/javascript/api/word/word.searchoptions#matchWholeWord)|获取或设置指示是否只查找整个单词，而不查找长单词的一部分的值。|
||[matchWildcards](/javascript/api/word/word.searchoptions#matchWildcards)|获取或设置指示搜索是否使用特殊搜索操作符执行的值。|
|[Section](/javascript/api/word/word.section)|[getFooter (类型：Word.HeaderFooterType) ](/javascript/api/word/word.section#getFooter_type_)|获取节的页脚之一。|
||[getHeader (类型：Word.HeaderFooterType) ](/javascript/api/word/word.section#getHeader_type_)|获取节的标头之一。|
||[body](/javascript/api/word/word.section#body)|获取节的 body 对象。|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Word JavaScript API 要求集](word-api-requirement-sets.md)

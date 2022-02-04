---
title: Word JavaScript API 要求集 1.1
description: 有关 WordApi 1.1 要求集的详细信息
ms.date: 11/01/2021
ms.prod: word
ms.localizationpriority: medium
---

# <a name="whats-new-in-word-javascript-api-11"></a>Word JavaScript API 1.1 的新增功能

WordApi 1.1 是 Word JavaScript API 的第一个要求集。 这是唯一受 Word API 要求集支持Word 2016。

## <a name="api-list"></a>API 列表

下表列出了 Word JavaScript API 要求集 1.1 中的 API。 若要查看受 Word JavaScript API 要求集 1.1 支持的所有 API 的 API 参考文档，请参阅要求集 [1.1 中的 Word API](/javascript/api/word?view=word-js-1.1&preserve-view=true)。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[正文](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#word-word-body-clear-member(1))|清除 body 对象的内容。|
||[contentControls](/javascript/api/word/word.body#word-word-body-contentcontrols-member)|获取正文中的格式文本内容控件对象的集合。|
||[font](/javascript/api/word/word.body#word-word-body-font-member)|获取正文的文本格式。|
||[getHtml()](/javascript/api/word/word.body#word-word-body-gethtml-member(1))|获取 body 对象的 HTML 表示形式。|
||[getOoxml()](/javascript/api/word/word.body#word-word-body-getooxml-member(1))|获取 body 对象的 OOXML (Office Open XML) 表示形式。|
||[ignorePunct](/javascript/api/word/word.body#word-word-body-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.body#word-word-body-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.body#word-word-body-inlinepictures-member)|获取正文中的 InlinePicture 对象的集合。|
||[insertBreak (breakType： Word.BreakType， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.body#word-word-body-insertbreak-member(1))|在主文档的指定位置插入分隔符。|
||[insertContentControl()](/javascript/api/word/word.body#word-word-body-insertcontentcontrol-member(1))|使用富文本内容控件封装 body 对象。|
||[insertFileFromBase64 (base64File： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.body#word-word-body-insertfilefrombase64-member(1))|将文档插入到正文中的指定位置。|
||[insertHtml (html： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.body#word-word-body-inserthtml-member(1))|在指定位置插入 HTML。|
||[insertOoxml (ooxml： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.body#word-word-body-insertooxml-member(1))|在指定位置插入 OOXML。|
||[insertParagraph (paragraphText： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.body#word-word-body-insertparagraph-member(1))|在指定位置插入段落。|
||[insertText (text： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.body#word-word-body-inserttext-member(1))|将文本插入到正文中的指定位置。|
||[matchCase](/javascript/api/word/word.body#word-word-body-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.body#word-word-body-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.body#word-word-body-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.body#word-word-body-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.body#word-word-body-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.body#word-word-body-paragraphs-member)|获取 body 中的 paragraph 对象的集合。|
||[parentContentControl](/javascript/api/word/word.body#word-word-body-parentcontentcontrol-member)|获取包含正文的内容控件。|
||[search (searchText： string， searchOptions？： Word.SearchOptions \| { ignorePunct？： boolean ignoreSpace？： boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： boolean matchWholeWord？： boolean matchWildcards？： boolean }) ](/javascript/api/word/word.body#word-word-body-search-member(1))|在 body 对象范围内使用指定的 SearchOptions 执行搜索。|
||[select (selectionMode？： Word.SelectionMode) ](/javascript/api/word/word.body#word-word-body-select-member(1))|选择正文并在 Word UI 中进行浏览。|
||[style](/javascript/api/word/word.body#word-word-body-style-member)|获取或设置正文的样式名称。|
||[text](/javascript/api/word/word.body#word-word-body-text-member)|获取正文的文本。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[appearance](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-appearance-member)|获取或设置内容控件的外观。|
||[cannotDelete](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-cannotdelete-member)|获取或设置指示用户是否可以删除内容控件的值。|
||[cannotEdit](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-cannotedit-member)|获取或设置指示用户是否可以编辑内容控件的内容的值。|
||[clear()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-clear-member(1))|清除内容控件的内容。|
||[color](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-color-member)|获取或设置内容控件的颜色。|
||[contentControls](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-contentcontrols-member)|获取内容控件中的内容控件对象的集合。|
||[delete (keepContent： boolean) ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-delete-member(1))|删除内容控件及其内容。|
||[font](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-font-member)|获取内容控件的文本格式。|
||[getHtml()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-gethtml-member(1))|获取内容控件对象的 HTML 表示形式。|
||[getOoxml()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getooxml-member(1))|获取内容控件对象的 Office Open XML (OOXML) 表示形式。|
||[id](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-id-member)|获取表示内容控件标识符的整数。|
||[ignorePunct](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inlinepictures-member)|获取内容控件中的 inlinePicture 对象的集合。|
||[insertBreak (breakType： Word.BreakType， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertbreak-member(1))|在主文档的指定位置插入分隔符。|
||[insertFileFromBase64 (base64File： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertfilefrombase64-member(1))|将文档插入到内容控件中的指定位置。|
||[insertHtml (html： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserthtml-member(1))|将 HTML 插入到内容控件中的指定位置。|
||[insertOoxml (ooxml： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertooxml-member(1))|将 OOXML 插入到位于指定位置的内容控件中。|
||[insertParagraph (paragraphText： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertparagraph-member(1))|在指定位置插入段落。|
||[insertText (text： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserttext-member(1))|将文本插入到内容控件中的指定位置。|
||[matchCase](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-paragraphs-member)|获取内容控件中的 paragraph 对象的集合。|
||[parentContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentcontentcontrol-member)|获取包含此内容控件的内容控件。|
||[placeholderText](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-placeholdertext-member)|获取或设置内容控件的占位符文本。|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-removewhenedited-member)|获取或设置指示内容控件在编辑后是否可以删除的值。|
||[search (searchText： string， searchOptions？： Word.SearchOptions \| { ignorePunct？： boolean ignoreSpace？： boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： boolean matchWholeWord？： boolean matchWildcards？： boolean }) ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-search-member(1))|在内容控件对象范围内使用指定的 SearchOptions 执行搜索。|
||[select (selectionMode？： Word.SelectionMode) ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-select-member(1))|选择内容控件。|
||[style](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-style-member)|获取或设置内容控件的样式名称。|
||[标记](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-tag-member)|获取或设置用于标识内容控件的标记。|
||[text](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-text-member)|获取内容控件的文本。|
||[title](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-title-member)|获取或设置内容控件的标题。|
||[type](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-type-member)|获取内容控件的类型。|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbyid-member(1))|按其标识符获取内容控件。|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytag-member(1))|获取具有指定标记的内容控件。|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytitle-member(1))|获取具有指定标题的内容控件。|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getitem-member(1))|按内容控件在集合中的索引获取内容控件。|
||[items](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-items-member)|获取此集合中已加载的子项。|
|[Document](/javascript/api/word/word.document)|[body](/javascript/api/word/word.document#word-word-document-body-member)|获取主文档的 body 对象。|
||[contentControls](/javascript/api/word/word.document#word-word-document-contentcontrols-member)|获取文档中的内容控件对象的集合。|
||[getSelection () ](/javascript/api/word/word.document#word-word-document-getselection-member(1))|获取文档的当前选定内容。|
||[save()](/javascript/api/word/word.document#word-word-document-save-member(1))|保存文档。|
||[saved](/javascript/api/word/word.document#word-word-document-saved-member)|指示是否已保存在文档中所做的更改。|
||[sections](/javascript/api/word/word.document#word-word-document-sections-member)|获取文档中 section 对象的集合。|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#word-word-font-bold-member)|获取或设置表示字体是否为粗体的值。|
||[color](/javascript/api/word/word.font#word-word-font-color-member)|获取或设置指定字体的颜色。|
||[doubleStrikeThrough](/javascript/api/word/word.font#word-word-font-doublestrikethrough-member)|获取或设置一个值，该值指示字体是否具有双删除线。|
||[highlightColor](/javascript/api/word/word.font#word-word-font-highlightcolor-member)|获取或设置突出显示颜色。|
||[italic](/javascript/api/word/word.font#word-word-font-italic-member)|获取或设置表示字体是否为斜体的值。|
||[name](/javascript/api/word/word.font#word-word-font-name-member)|获取或设置表示字体名称的值。|
||[size](/javascript/api/word/word.font#word-word-font-size-member)|获取或设置表示字体大小（以磅值表示）的值。|
||[strikeThrough](/javascript/api/word/word.font#word-word-font-strikethrough-member)|获取或设置一个值，该值指示字体是否有删除线。|
||[subscript](/javascript/api/word/word.font#word-word-font-subscript-member)|获取或设置表示字体是否为下标的值。|
||[superscript](/javascript/api/word/word.font#word-word-font-superscript-member)|获取或设置表示字体是否为上标的值。|
||[underline](/javascript/api/word/word.font#word-word-font-underline-member)|获取或设置表示字体的下划线类型的值。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-alttextdescription-member)|获取或设置一个字符串，该字符串代表与内联图像关联的可选文本。|
||[altTextTitle](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-alttexttitle-member)|获取或设置包含嵌入式图像的标题的字符串。|
||[getBase64ImageSrc () ](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getbase64imagesrc-member(1))|获取嵌入式图像的 base64 编码的字符串表示。|
||[height](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-height-member)|获取或设置描述嵌入式图像的高度的数字。|
||[hyperlink](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-hyperlink-member)|获取或设置图像上的超链接。|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertcontentcontrol-member(1))|使用富文本内容控件封装嵌入式图像。|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-lockaspectratio-member)|获取或设置指示在您调整嵌入式图像大小时其是否保留原始比例的值。|
||[parentContentControl](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentcontentcontrol-member)|获取包含嵌入式图像的内容控件。|
||[width](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-width-member)|获取或设置描述嵌入式图像的宽度的数字。|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-items-member)|获取此集合中已加载的子项。|
|[Paragraph](/javascript/api/word/word.paragraph)|[alignment](/javascript/api/word/word.paragraph#word-word-paragraph-alignment-member)|获取或设置段落的对齐方式。|
||[clear()](/javascript/api/word/word.paragraph#word-word-paragraph-clear-member(1))|清除 paragraph 对象的内容。|
||[contentControls](/javascript/api/word/word.paragraph#word-word-paragraph-contentcontrols-member)|获取段落中的内容控件对象的集合。|
||[delete()](/javascript/api/word/word.paragraph#word-word-paragraph-delete-member(1))|从文档中删除段落及其内容。|
||[firstLineIndent](/javascript/api/word/word.paragraph#word-word-paragraph-firstlineindent-member)|获取或设置首行缩进或悬挂缩进的大小（以磅值表示）。|
||[font](/javascript/api/word/word.paragraph#word-word-paragraph-font-member)|获取段落的文本格式。|
||[getHtml()](/javascript/api/word/word.paragraph#word-word-paragraph-gethtml-member(1))|获取 paragraph 对象的 HTML 表示形式。|
||[getOoxml()](/javascript/api/word/word.paragraph#word-word-paragraph-getooxml-member(1))|获取 paragraph 对象的 Office Open XML (OOXML) 表示形式。|
||[ignorePunct](/javascript/api/word/word.paragraph#word-word-paragraph-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.paragraph#word-word-paragraph-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.paragraph#word-word-paragraph-inlinepictures-member)|获取段落中的 InlinePicture 对象的集合。|
||[insertBreak (breakType： Word.BreakType， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.paragraph#word-word-paragraph-insertbreak-member(1))|在主文档的指定位置插入分隔符。|
||[insertContentControl()](/javascript/api/word/word.paragraph#word-word-paragraph-insertcontentcontrol-member(1))|使用富文本内容控件封装 paragraph 对象。|
||[insertFileFromBase64 (base64File： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.paragraph#word-word-paragraph-insertfilefrombase64-member(1))|将文档插入到段落中的指定位置。|
||[insertHtml (html： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.paragraph#word-word-paragraph-inserthtml-member(1))|将 HTML 插入到段落中的指定位置。|
||[insertInlinePictureFromBase64 (base64EncodedImage： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.paragraph#word-word-paragraph-insertinlinepicturefrombase64-member(1))|将图片插入到段落中的指定位置。|
||[insertOoxml (ooxml： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.paragraph#word-word-paragraph-insertooxml-member(1))|将 OOXML 插入到段落中的指定位置。|
||[insertParagraph (paragraphText： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.paragraph#word-word-paragraph-insertparagraph-member(1))|在指定位置插入段落。|
||[insertText (text： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.paragraph#word-word-paragraph-inserttext-member(1))|将文本插入到段落中的指定位置。|
||[leftIndent](/javascript/api/word/word.paragraph#word-word-paragraph-leftindent-member)|获取或设置段落的向左缩进值（以磅值表示）。|
||[lineSpacing](/javascript/api/word/word.paragraph#word-word-paragraph-linespacing-member)|获取或设置指定段落的行间距（以磅值表示）。|
||[lineUnitAfter](/javascript/api/word/word.paragraph#word-word-paragraph-lineunitafter-member)|获取或设置段落后的间距，以网格线表示。|
||[lineUnitBefore](/javascript/api/word/word.paragraph#word-word-paragraph-lineunitbefore-member)|获取或设置段落前面的网格线中的间隔量。|
||[matchCase](/javascript/api/word/word.paragraph#word-word-paragraph-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.paragraph#word-word-paragraph-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.paragraph#word-word-paragraph-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.paragraph#word-word-paragraph-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.paragraph#word-word-paragraph-matchwildcards-member)||
||[outlineLevel](/javascript/api/word/word.paragraph#word-word-paragraph-outlinelevel-member)|获取或设置段落的大纲级别。|
||[parentContentControl](/javascript/api/word/word.paragraph#word-word-paragraph-parentcontentcontrol-member)|获取包含段落的内容控件。|
||[rightIndent](/javascript/api/word/word.paragraph#word-word-paragraph-rightindent-member)|获取或设置段落的向右缩进值（以磅值表示）。|
||[search (searchText： string， searchOptions？： Word.SearchOptions \| { ignorePunct？： boolean ignoreSpace？： boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： boolean matchWholeWord？： boolean matchWildcards？： boolean }) ](/javascript/api/word/word.paragraph#word-word-paragraph-search-member(1))|使用指定的 SearchOptions 搜索段落对象的范围。|
||[select (selectionMode？： Word.SelectionMode) ](/javascript/api/word/word.paragraph#word-word-paragraph-select-member(1))|选择并在 Word UI 中导航到段落。|
||[spaceAfter](/javascript/api/word/word.paragraph#word-word-paragraph-spaceafter-member)|获取或设置段落后面的间距（以磅值表示）。|
||[spaceBefore](/javascript/api/word/word.paragraph#word-word-paragraph-spacebefore-member)|获取或设置段落前面的间距（以磅值表示）。|
||[style](/javascript/api/word/word.paragraph#word-word-paragraph-style-member)|获取或设置段落的样式名称。|
||[text](/javascript/api/word/word.paragraph#word-word-paragraph-text-member)|获取段落的文本。|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-items-member)|获取此集合中已加载的子项。|
|[区域](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#word-word-range-clear-member(1))|清除 range 对象的内容。|
||[contentControls](/javascript/api/word/word.range#word-word-range-contentcontrols-member)|获取范围中的内容控件对象的集合。|
||[delete()](/javascript/api/word/word.range#word-word-range-delete-member(1))|从文档中删除区域及其内容。|
||[font](/javascript/api/word/word.range#word-word-range-font-member)|获取区域的文本格式。|
||[getHtml()](/javascript/api/word/word.range#word-word-range-gethtml-member(1))|获取 range 对象的 HTML 表示形式。|
||[getOoxml()](/javascript/api/word/word.range#word-word-range-getooxml-member(1))|获取 range 对象的 OOXML 表示形式。|
||[ignorePunct](/javascript/api/word/word.range#word-word-range-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.range#word-word-range-ignorespace-member)||
||[insertBreak (breakType： Word.BreakType， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.range#word-word-range-insertbreak-member(1))|在主文档的指定位置插入分隔符。|
||[insertContentControl()](/javascript/api/word/word.range#word-word-range-insertcontentcontrol-member(1))|使用富文本内容控件封装 range 对象。|
||[insertFileFromBase64 (base64File： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.range#word-word-range-insertfilefrombase64-member(1))|在指定位置插入 document。|
||[insertHtml (html： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.range#word-word-range-inserthtml-member(1))|在指定位置插入 HTML。|
||[insertOoxml (ooxml： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.range#word-word-range-insertooxml-member(1))|在指定位置插入 OOXML。|
||[insertParagraph (paragraphText： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.range#word-word-range-insertparagraph-member(1))|在指定位置插入段落。|
||[insertText (text： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.range#word-word-range-inserttext-member(1))|在指定位置插入文本。|
||[matchCase](/javascript/api/word/word.range#word-word-range-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.range#word-word-range-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.range#word-word-range-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.range#word-word-range-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.range#word-word-range-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.range#word-word-range-paragraphs-member)|获取 range 中的 paragraph 对象的集合。|
||[parentContentControl](/javascript/api/word/word.range#word-word-range-parentcontentcontrol-member)|获取包含该范围的内容控件。|
||[search (searchText： string， searchOptions？： Word.SearchOptions \| { ignorePunct？： boolean ignoreSpace？： boolean matchCase？： boolean matchPrefix？： boolean matchSuffix？： boolean matchWholeWord？： boolean matchWildcards？： boolean }) ](/javascript/api/word/word.range#word-word-range-search-member(1))|使用指定的 SearchOptions 搜索 range 对象的范围。|
||[select (selectionMode？： Word.SelectionMode) ](/javascript/api/word/word.range#word-word-range-select-member(1))|选择并在 Word UI 中导航到区域。|
||[style](/javascript/api/word/word.range#word-word-range-style-member)|获取或设置范围的样式名称。|
||[text](/javascript/api/word/word.range#word-word-range-text-member)|获取区域的文本。|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#word-word-rangecollection-items-member)|获取此集合中已加载的子项。|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#word-word-searchoptions-ignorepunct-member)|获取或设置指示是否忽略单词之间的所有标点符号的值。|
||[ignoreSpace](/javascript/api/word/word.searchoptions#word-word-searchoptions-ignorespace-member)|获取或设置一个值，该值指示是否忽略单词之间的所有空格。|
||[matchCase](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchcase-member)|获取或设置指示是否执行区分大小写的搜索的值。|
||[matchPrefix](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchprefix-member)|获取或设置指示是否匹配以搜索字符串开头的单词。|
||[matchSuffix](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchsuffix-member)|获取或设置指示是否匹配以搜索字符串结尾的单词。|
||[matchWholeWord](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchwholeword-member)|获取或设置指示是否只查找整个单词，而不查找长单词的一部分的值。|
||[matchWildcards](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchwildcards-member)|获取或设置指示搜索是否使用特殊搜索操作符执行的值。|
|[Section](/javascript/api/word/word.section)|[body](/javascript/api/word/word.section#word-word-section-body-member)|获取节的 body 对象。|
||[getFooter (类型：Word.HeaderFooterType) ](/javascript/api/word/word.section#word-word-section-getfooter-member(1))|获取节的页脚之一。|
||[getHeader (类型：Word.HeaderFooterType) ](/javascript/api/word/word.section#word-word-section-getheader-member(1))|获取节的标头之一。|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-items-member)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Word JavaScript API 要求集](word-api-requirement-sets.md)

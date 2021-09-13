---
title: Word JavaScript API 要求集 1.2
description: 有关 WordApi 1.2 要求集的详细信息
ms.date: 11/09/2020
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: de293cf67bbb452fe3c2b8c5de4896adf5cf7a43
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152633"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Word JavaScript API 1.2 的最近更新

WordApi 1.2 增加了对内联图片的支持。

## <a name="api-list"></a>API 列表

下表列出了 Word JavaScript API 要求集 1.2 中的 API。 若要查看受 Word JavaScript API 要求集 1.2 或更早版本支持的所有 API 的 API 参考文档，请参阅要求集 [1.2](/javascript/api/word?view=word-js-1.2&preserve-view=true)或更早版本中的 Word API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[正文](/javascript/api/word/word.body)|[insertInlinePictureFromBase64 (base64EncodedImage： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.body#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|将图片插入到正文中的指定位置。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64 (base64EncodedImage： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.contentcontrol#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|将嵌入式图片插入到内容控件中的指定位置。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#delete__)|从文档中删除嵌入式图片。|
||[insertBreak (breakType： Word.BreakType， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.inlinepicture#insertBreak_breakType__insertLocation_)|在主文档的指定位置插入分隔符。|
||[insertFileFromBase64 (base64File： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.inlinepicture#insertFileFromBase64_base64File__insertLocation_)|在指定位置插入 document。|
||[insertHtml (html： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.inlinepicture#insertHtml_html__insertLocation_)|在指定位置插入 HTML。|
||[insertInlinePictureFromBase64 (base64EncodedImage： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.inlinepicture#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|在指定位置插入 inlinePicture。|
||[insertOoxml (ooxml： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.inlinepicture#insertOoxml_ooxml__insertLocation_)|在指定位置插入 OOXML。|
||[insertParagraph (paragraphText： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.inlinepicture#insertParagraph_paragraphText__insertLocation_)|在指定位置插入段落。|
||[insertText (text： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.inlinepicture#insertText_text__insertLocation_)|在指定位置插入文本。|
||[paragraph](/javascript/api/word/word.inlinepicture#paragraph)|获取包含嵌入式图像的父段落。|
||[select (selectionMode？： Word.SelectionMode) ](/javascript/api/word/word.inlinepicture#select_selectionMode_)|选择 inlinePicture。|
|[Range](/javascript/api/word/word.range)|[insertInlinePictureFromBase64 (base64EncodedImage： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.range#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|在指定位置插入图片。|
||[inlinePictures](/javascript/api/word/word.range#inlinePictures)|获取 range 中的一组 inlinePicture 对象。|

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Word JavaScript API 要求集](word-api-requirement-sets.md)

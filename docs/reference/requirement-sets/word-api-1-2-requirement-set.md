---
title: Word JavaScript API 要求集 1.2
description: 有关 WordApi 1.2 要求集的详细信息。
ms.date: 11/09/2020
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: 1a5af83615786b241c43ecb07ee0d23b3758cfc8
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744217"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Word JavaScript API 1.2 的最近更新

WordApi 1.2 增加了对内联图片的支持。

## <a name="api-list"></a>API 列表

下表列出了 Word JavaScript API 要求集 1.2 中的 API。 若要查看受 Word JavaScript API 要求集 1.2 或更早版本支持的所有 API 的 API 参考文档，请参阅要求集 [1.2 或更早版本中的 Word API](/javascript/api/word?view=word-js-1.2&preserve-view=true)。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[正文](/javascript/api/word/word.body)|[insertInlinePictureFromBase64 (base64EncodedImage： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.body#word-word-body-insertinlinepicturefrombase64-member(1))|将图片插入到正文中的指定位置。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64 (base64EncodedImage： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertinlinepicturefrombase64-member(1))|将嵌入式图片插入到内容控件中的指定位置。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-delete-member(1))|从文档中删除嵌入式图片。|
||[insertBreak (breakType： Word.BreakType， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertbreak-member(1))|在主文档的指定位置插入分隔符。|
||[insertFileFromBase64 (base64File： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertfilefrombase64-member(1))|在指定位置插入 document。|
||[insertHtml (html： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-inserthtml-member(1))|在指定位置插入 HTML。|
||[insertInlinePictureFromBase64 (base64EncodedImage： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertinlinepicturefrombase64-member(1))|在指定位置插入 inlinePicture。|
||[insertOoxml (ooxml： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertooxml-member(1))|在指定位置插入 OOXML。|
||[insertParagraph (paragraphText： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertparagraph-member(1))|在指定位置插入段落。|
||[insertText (text： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-inserttext-member(1))|在指定位置插入文本。|
||[paragraph](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-paragraph-member)|获取包含嵌入式图像的父段落。|
||[select (selectionMode？： Word.SelectionMode) ](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-select-member(1))|选择 inlinePicture。|
|[范围](/javascript/api/word/word.range)|[inlinePictures](/javascript/api/word/word.range#word-word-range-inlinepictures-member)|获取 range 中的一组 inlinePicture 对象。|
||[insertInlinePictureFromBase64 (base64EncodedImage： string， insertLocation： Word.InsertLocation) ](/javascript/api/word/word.range#word-word-range-insertinlinepicturefrombase64-member(1))|在指定位置插入图片。|

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Word JavaScript API 要求集](word-api-requirement-sets.md)

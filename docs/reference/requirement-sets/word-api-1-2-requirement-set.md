---
title: Word JavaScript API 要求集1。2
description: 有关 WordApi 1.2 要求集的详细信息
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: ee9bf60a3a944a3a01a2ca5aa10d01958e3d3475
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996422"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Word JavaScript API 1.2 的最近更新

WordApi 1.2 增加了对内嵌图片的支持。

## <a name="api-list"></a>API 列表

下表列出了 Word JavaScript API 要求集1.2 中的 Api。 若要查看 Word JavaScript API 要求集1.2 或更早版本支持的所有 Api 的 API 参考文档，请参阅 [要求集1.2 或更早版本中的 Word api](/javascript/api/word?view=word-js-1.2&preserve-view=true)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[正文](/javascript/api/word/word.body)|[insertInlinePictureFromBase64 (base64EncodedImage： string，insertLocation： InsertLocation) ](/javascript/api/word/word.body#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|将图片插入到正文中的指定位置。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64 (base64EncodedImage： string，insertLocation： InsertLocation) ](/javascript/api/word/word.contentcontrol#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|将嵌入式图片插入到内容控件中的指定位置。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#delete--)|从文档中删除嵌入式图片。|
||[insertBreak (breakType： BreakType，insertLocation： Word. InsertLocation) ](/javascript/api/word/word.inlinepicture#insertbreak-breaktype--insertlocation-)|在主文档的指定位置插入分隔符。|
||[insertFileFromBase64 (base64File： string，insertLocation： InsertLocation) ](/javascript/api/word/word.inlinepicture#insertfilefrombase64-base64file--insertlocation-)|在指定位置插入 document。|
||[insertHtml (html： string，insertLocation： InsertLocation) ](/javascript/api/word/word.inlinepicture#inserthtml-html--insertlocation-)|在指定位置插入 HTML。|
||[insertInlinePictureFromBase64 (base64EncodedImage： string，insertLocation： InsertLocation) ](/javascript/api/word/word.inlinepicture#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|在指定位置插入 inlinePicture。|
||[insertOoxml (ooxml： string，insertLocation： InsertLocation) ](/javascript/api/word/word.inlinepicture#insertooxml-ooxml--insertlocation-)|在指定位置插入 OOXML。|
||[insertParagraph (paragraphText： string，insertLocation： InsertLocation) ](/javascript/api/word/word.inlinepicture#insertparagraph-paragraphtext--insertlocation-)|在指定位置插入段落。|
||[insertText (text： string，insertLocation： InsertLocation) ](/javascript/api/word/word.inlinepicture#inserttext-text--insertlocation-)|在指定位置插入文本。|
||[paragraph](/javascript/api/word/word.inlinepicture#paragraph)|获取包含嵌入式图像的父段落。|
||[选择 (selectionMode？： SelectionMode) ](/javascript/api/word/word.inlinepicture#select-selectionmode-)|选择 inlinePicture。|
|[Range](/javascript/api/word/word.range)|[insertInlinePictureFromBase64 (base64EncodedImage： string，insertLocation： InsertLocation) ](/javascript/api/word/word.range#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|在指定位置插入图片。|
||[inlinePictures](/javascript/api/word/word.range#inlinepictures)|获取 range 中的一组 inlinePicture 对象。|

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Word JavaScript API 要求集](word-api-requirement-sets.md)

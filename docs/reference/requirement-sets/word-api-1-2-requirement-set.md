---
title: Word JavaScript API 要求集1。2
description: 有关 WordApi 1.2 要求集的详细信息
ms.date: 07/17/2019
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: c6244b7ce9ff7b5cbde68baad26e60a6326199d8
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/19/2019
ms.locfileid: "35804704"
---
# <a name="whats-new-in-word-javascript-api-12"></a>Word JavaScript API 1.2 的最近更新

WordApi 1.2 增加了对内嵌图片的支持。

## <a name="api-list"></a>API 列表

下表列出了作为 WordApi 1.2 要求集的一部分添加的 Api。

| Class | 域 | 说明 |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[insertInlinePictureFromBase64 (base64EncodedImage: string, insertLocation: InsertLocation)](/javascript/api/word/word.body#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|将图片插入到正文中的指定位置。 insertLocation 值可以为“Start”或“End”。|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[insertInlinePictureFromBase64 (base64EncodedImage: string, insertLocation: InsertLocation)](/javascript/api/word/word.contentcontrol#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|将嵌入式图片插入到内容控件中的指定位置。 insertLocation 值可以为“Replace”、“Start”或“End”。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[delete()](/javascript/api/word/word.inlinepicture#delete--)|从文档中删除嵌入式图片。|
||[insertBreak (breakType: BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.inlinepicture#insertbreak-breaktype--insertlocation-)|在主文档的指定位置插入分隔符。 insertLocation 值可以为“Before”或“After”。|
||[insertFileFromBase64 (base64File: string, insertLocation: InsertLocation)](/javascript/api/word/word.inlinepicture#insertfilefrombase64-base64file--insertlocation-)|在指定位置插入 document。 insertLocation 值可以为“Before”或“After”。|
||[insertHtml (html: string, insertLocation: InsertLocation)](/javascript/api/word/word.inlinepicture#inserthtml-html--insertlocation-)|在指定位置插入 HTML。 insertLocation 值可以为“Before”或“After”。|
||[insertInlinePictureFromBase64 (base64EncodedImage: string, insertLocation: InsertLocation)](/javascript/api/word/word.inlinepicture#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|在指定位置插入 inlinePicture。 InsertLocation 值可以是 "Replace"、"Before" 或 "After"。|
||[insertOoxml (ooxml: string, insertLocation: InsertLocation)](/javascript/api/word/word.inlinepicture#insertooxml-ooxml--insertlocation-)|在指定位置插入 OOXML。  insertLocation 值可以为“Before”或“After”。|
||[insertParagraph (paragraphText: string, insertLocation: InsertLocation)](/javascript/api/word/word.inlinepicture#insertparagraph-paragraphtext--insertlocation-)|在指定位置插入段落。 insertLocation 值可以为“Before”或“After”。|
||[insertText (text: string, insertLocation: InsertLocation)](/javascript/api/word/word.inlinepicture#inserttext-text--insertlocation-)|在指定位置插入文本。 insertLocation 的可取值为“Before”或“After”。|
||[paragraph](/javascript/api/word/word.inlinepicture#paragraph)|获取包含嵌入式图像的父段落。 只读。|
||[select (selectionMode？: SelectionMode)](/javascript/api/word/word.inlinepicture#select-selectionmode-)|选择 inlinePicture。 这会导致 Word 滚动到选定内容。|
|[Range](/javascript/api/word/word.range)|[insertInlinePictureFromBase64 (base64EncodedImage: string, insertLocation: InsertLocation)](/javascript/api/word/word.range#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|在指定位置插入图片。 InsertLocation 值可以是 "Replace"、"Start"、"End"、"Before" 或 "After"。|
||[inlinePictures](/javascript/api/word/word.range#inlinepictures)|获取 range 中的一组 inlinePicture 对象。 只读。|

## <a name="see-also"></a>另请参阅

- [Word JavaScript API 参考文档](/javascript/api/word)
- [Word JavaScript API 要求集](word-api-requirement-sets.md)

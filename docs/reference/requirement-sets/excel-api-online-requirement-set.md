---
title: Excel JavaScript API 仅联机要求集
description: 有关 ExcelApiOnline 要求集的详细信息
ms.date: 12/05/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad2a3cd627552baeb449397fa917fe10e86ebbaf
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814150"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel JavaScript API 仅联机要求集

`ExcelApiOnline`要求集是一个特殊要求集，其中包含仅适用于 web 上的 Excel 的功能。 此要求集中的 Api 被认为是针对 web 主机上的 Excel 的生产 Api （不受未记录的行为或结构更改）。 `ExcelApiOnline`被视为针对其他平台（Windows、Mac、iOS）的 "预览" Api，这些平台可能不支持这些平台。

当在所有平台`ExcelApiOnline`上支持要求集中的 api 时，它们将添加到下一个发布的要求集`ExcelApi 1.[NEXT]`（）。 一旦新要求是公共的，将从这些 Api 中`ExcelApiOnline`删除。 可将此视为将 API 从预览迁移到发布的类似升级过程。

> [!IMPORTANT]
> `ExcelApiOnline`是最新编号的要求集的超集。

> [!IMPORTANT]
> `ExcelApiOnline 1.1`是仅联机 Api 的唯一版本。 这是因为 web 上的 Excel 将始终有一个版本可供最新版本的用户使用。

## <a name="recommended-usage"></a>建议使用

由于`ExcelApiOnline` web 上的 Excel 仅支持 api，因此，您的外接程序应检查是否支持要求集，然后再调用这些 api。 这样可以避免在不同的平台上调用仅联机 API。

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

一旦 API 位于跨平台要求集，就应删除或编辑该`isSetSupported`检查。 这将在其他平台上启用外接程序的功能。 进行此更改时，请务必在这些平台上测试功能。

> [!IMPORTANT]
> 清单不能指定`ExcelApiOnline 1.1`为激活要求。 不是在[Set 元素](../manifest/set.md)中使用的有效值。

## <a name="api-list"></a>API 列表

下面的 Api 当前可用于 web 上的 Excel，作为`ExcelApiOnline 1.1`要求集的一部分。

| Class | 域 | 说明 |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[提及](/javascript/api/excel/excel.comment#mentions)|获取注释中提到的实体（如人员）。|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|获取丰富的注释内容（例如，注释中的提及）。 此字符串不应显示给最终用户。 您的外接程序应仅使用此信息分析丰富的注释内容。|
||[updateMentions （contentWithMentions： CommentRichContent）](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|使用特殊格式的字符串和提及列表更新注释内容。|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|获取或设置注释中提到的实体的电子邮件地址。|
||[id](/javascript/api/excel/excel.commentmention#id)|获取或设置实体的 id。 这与中`CommentRichContent.richContent`的一个 id 相匹配。|
||[name](/javascript/api/excel/excel.commentmention#name)|获取或设置注释中提到的实体的名称。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[提及](/javascript/api/excel/excel.commentreply#mentions)|获取注释中提到的实体（如人员）。|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|获取丰富的注释内容（例如，注释中的提及）。 此字符串不应显示给最终用户。 您的外接程序应仅使用此信息分析丰富的注释内容。|
||[updateMentions （contentWithMentions： CommentRichContent）](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|使用特殊格式的字符串和提及列表更新注释内容。|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[提及](/javascript/api/excel/excel.commentrichcontent#mentions)|包含注释中提到的所有实体（例如，人员）的数组。|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)||
|[Range](/javascript/api/excel/excel.range)|[moveTo （destinationRange： Range \|字符串）](/javascript/api/excel/excel.range#moveto-destinationrange-)|将单元格的值、格式和公式从当前区域移动到目标区域，替换这些单元格中的旧信息。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent （金额：数字）](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|调整范围格式的缩进量。 缩进值的范围为0到250，以字符为单位。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-online)
- [Excel JavaScript 预览 API](./excel-preview-apis.md)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
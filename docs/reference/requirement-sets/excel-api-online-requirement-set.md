---
title: Excel JavaScript API 仅联机要求集
description: 有关 ExcelApiOnline 要求集的详细信息。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 282e11e415d51a6724715091d894df64ebaabfae
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604678"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel JavaScript API 仅联机要求集

要求 `ExcelApiOnline` 集是一个特殊要求集，其中包括仅适用于 Excel 网页应用的功能。 此要求集的 API 被视为生产 API， (Web 应用程序上的 Excel) 记录的行为或结构更改。 `ExcelApiOnline` API 被视为 Windows、Mac、iOS (其他平台的"预览"API) 并且可能不受这些平台的任何支持。

当所有平台都支持要求集内 API 时，它们将被添加到下一个发布的要求 `ExcelApiOnline` `ExcelApi 1.[NEXT]` () 。 一旦该新要求公开，将从 中删除这些 `ExcelApiOnline` API。 将此过程视为从预览版移动到发布的 API 的类似推广过程。

> [!IMPORTANT]
> `ExcelApiOnline` 是最新编号要求集的超集。

> [!IMPORTANT]
> `ExcelApiOnline 1.1` 是仅联机 API 的唯一版本。 这是因为 Excel 网页版将始终有一个版本可供最新版本的用户使用。

下表提供了 API 的简要摘要，而后续 API 列表表提供了当前 [API](#api-list) `ExcelApiOnline` 的详细列表。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 命名工作表视图 | 以编程方式控制每用户工作表视图。 | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |

## <a name="recommended-usage"></a>建议的用法

由于 API 仅受 Web 上的 Excel 支持，因此加载项应在调用这些 API 之前检查 `ExcelApiOnline` 要求集是否受支持。 这可以避免在不同的平台上调用仅联机 API。

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

API 位于跨平台要求集后，应删除或编辑 `isSetSupported` 检查。 这将在其他平台上启用外接程序的功能。 进行此更改时，请务必在这些平台上测试该功能。

> [!IMPORTANT]
> 清单不能指定为 `ExcelApiOnline 1.1` 激活要求。 它在 Set 元素中不是有效的 [值](../manifest/set.md)。

## <a name="api-list"></a>API 列表

下表列出了要求集当前包含的 Excel JavaScript `ExcelApiOnline` API。 有关所有 Excel JavaScript API 的完整列表， (API 和以前发布的 API `ExcelApiOnline`) ，请参阅所有[Excel JavaScript API。](/javascript/api/excel?view=excel-js-online&preserve-view=true)

| 类 | 域 | 说明 |
|:---|:---|:---|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|激活此工作表视图。|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|从工作表中删除工作表视图。|
||[duplicate (name？： string) ](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|创建此工作表视图的副本。|
||[name](/javascript/api/excel/excel.namedsheetview#name)|获取或设置工作表视图的名称。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|创建具有给定名称的新工作表视图。|
||[enterTemporary () ](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|创建并激活新的临时工作表视图。|
||[exit () ](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|退出当前活动的工作表视图。|
||[getActive () ](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|获取工作表当前的活动工作表视图。|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|获取此工作表中的工作表视图数。|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|使用工作表视图的名称获取工作表视图。|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|按工作表视图在集合中的索引获取工作表视图。|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|获取此集合中已加载的子项。|
|[Range](/javascript/api/excel/excel.range)|[getExtendedRange (方向： Excel.KeyboardDirection， activeCell？： Range \| string) ](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|返回一个 range 对象，该对象包括当前区域以及区域边缘，根据提供的方向。|
||[getMergedAreas () ](/javascript/api/excel/excel.range#getmergedareas--)|返回 `RangeAreas` 一个对象，该对象代表此范围中的合并区域。|
||[getRangeEdge (方向： Excel.KeyboardDirection， activeCell？： Range \| string) ](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|返回一个 range 对象，该对象是数据区域的边缘单元格，对应于提供的方向。|
|[Table](/javascript/api/excel/excel.table)|[resize (newRange：Range \| string) ](/javascript/api/excel/excel.table#resize-newrange-)|将表格调整到新区域。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|返回工作表中呈现的工作表视图的集合。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Excel JavaScript 预览 API](excel-preview-apis.md)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)

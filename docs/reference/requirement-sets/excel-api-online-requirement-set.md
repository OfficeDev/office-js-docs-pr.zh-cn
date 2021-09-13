---
title: ExcelJavaScript API 仅联机要求集
description: 有关 ExcelApiOnline 要求集的详细信息。
ms.date: 07/23/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 067bd3f033c415b0a6ac271aa1f132cfeb92730c
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152241"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>ExcelJavaScript API 仅联机要求集

要求 `ExcelApiOnline` 集是一个特殊要求集，其中包括仅适用于Excel web 版。 此要求集内 API 被视为生产 API， (应用程序未记录的行为或) 更改Excel web 版 API。 `ExcelApiOnline`API 被视为适用于其他平台（如 (Windows、Mac、iOS) ）的"预览"API，可能不受这些平台的任何支持。

当要求集内 API 在所有平台中均受支持时，它们将被添加到下一个发布的要求 `ExcelApiOnline` `ExcelApi 1.[NEXT]` () 。 一旦该新要求公开，将从 中删除这些 `ExcelApiOnline` API。 将此过程视为与从预览版移动到发布的 API 类似的推广过程。

> [!IMPORTANT]
> `ExcelApiOnline` 是最新编号要求集的超集。

> [!IMPORTANT]
> `ExcelApiOnline 1.1` 是仅联机 API 的唯一版本。 这是因为Excel web 版版本始终为最新版本的用户可用。

下表提供了 API 的简要摘要，而后续 API 列表表提供了当前 [API](#api-list) `ExcelApiOnline` 的详细列表。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 命名工作表视图 | 以编程方式控制每用户工作表视图。 | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |

## <a name="recommended-usage"></a>建议的用法

由于 API 仅受 Excel web 版，因此加载项应在调用这些 API 之前检查 `ExcelApiOnline` 要求集是否受支持。 这可以避免在不同的平台上调用仅联机 API。

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

API 位于跨平台要求集后，应删除或编辑 `isSetSupported` 检查。 这将在其他平台上启用外接程序的功能。 进行此更改时，请务必在这些平台上测试该功能。

> [!IMPORTANT]
> 清单不能指定为 `ExcelApiOnline 1.1` 激活要求。 它在 Set 元素中不是有效的 [值](../manifest/set.md)。

## <a name="api-list"></a>API 列表

下表列出了要求Excel当前包含的 JavaScript `ExcelApiOnline` API。 有关所有 JavaScript EXCEL的完整列表 (包括 API 和以前发布的 API) ，请参阅所有 Excel `ExcelApiOnline` [JavaScript API。](/javascript/api/excel?view=excel-js-online&preserve-view=true)

| 类 | 域 | 说明 |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria (columnIndex： number) ](/javascript/api/excel/excel.autofilter#clearColumnCriteria_columnIndex_)|清除自动筛选的列筛选条件。|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate__)|激活此工作表视图。|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete__)|从工作表中删除工作表视图。|
||[duplicate (name？： string) ](/javascript/api/excel/excel.namedsheetview#duplicate_name_)|创建此工作表视图的副本。|
||[name](/javascript/api/excel/excel.namedsheetview#name)|获取或设置工作表视图的名称。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add_name_)|创建具有给定名称的新工作表视图。|
||[enterTemporary () ](/javascript/api/excel/excel.namedsheetviewcollection#enterTemporary__)|创建并激活新的临时工作表视图。|
||[exit () ](/javascript/api/excel/excel.namedsheetviewcollection#exit__)|退出当前活动的工作表视图。|
||[getActive () ](/javascript/api/excel/excel.namedsheetviewcollection#getActive__)|获取工作表当前的活动工作表视图。|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getCount__)|获取此工作表中的工作表视图数。|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItem_key_)|使用工作表视图的名称获取工作表视图。|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getItemAt_index_)|按工作表视图在集合中的索引获取工作表视图。|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|获取此集合中已加载的子项。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedSheetViews)|返回工作表中呈现的工作表视图的集合。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Excel JavaScript 预览 API](excel-preview-apis.md)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)

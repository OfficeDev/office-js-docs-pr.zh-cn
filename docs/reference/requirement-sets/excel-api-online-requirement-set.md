---
title: ExcelJavaScript API 仅联机要求集
description: 有关 ExcelApiOnline 要求集的详细信息。
ms.date: 09/16/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 9b8d326e1a756a873fc19b3d78f795ebf04e5f4e
ms.sourcegitcommit: a854a2fd2ad9f379a3ef712f307e0b1bb9b5b00d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/22/2021
ms.locfileid: "59474334"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>ExcelJavaScript API 仅联机要求集

要求 `ExcelApiOnline` 集是一个特殊要求集，其中包含仅适用于Excel web 版。 此要求集内 API 被视为生产 API， (应用程序的行为或结构) 未Excel web 版更改。 `ExcelApiOnline`API 被视为适用于其他平台（如 (Windows、Mac、iOS) ）的"预览"API，可能不受这些平台的任何支持。

当要求集内 API 在所有平台中均受支持时，它们将被添加到下一个发布的要求 `ExcelApiOnline` `ExcelApi 1.[NEXT]` () 。 一旦该新要求公开，将从 中删除这些 `ExcelApiOnline` API。 将此过程视为与从预览版移动到发布的 API 类似的推广过程。

> [!IMPORTANT]
> `ExcelApiOnline` 是最新编号要求集的超集。

> [!IMPORTANT]
> `ExcelApiOnline 1.1` 是仅联机 API 的唯一版本。 这是因为Excel web 版版本始终为最新版本的用户可用。

下表提供了 API 的简要摘要，而后续 API 列表表提供了当前 [API](#api-list) `ExcelApiOnline` 的详细列表。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 链接的工作簿 | 管理工作簿之间的链接，包括对刷新和断开工作簿链接的支持。 | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook) [、LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| 命名工作表视图 | 以编程方式控制每用户工作表视图。 | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) [、NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection) |

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

下表列出了要求Excel中当前包含的 JavaScript `ExcelApiOnline` API。 有关所有 JavaScript EXCEL的完整列表 (包括 API 和以前发布的 API `ExcelApiOnline`) ，请参阅所有 Excel [JavaScript API。](/javascript/api/excel?view=excel-js-online&preserve-view=true)

| 类 | 域 | 说明 |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria (columnIndex： number) ](/javascript/api/excel/excel.autofilter#clearColumnCriteria_columnIndex_)|清除自动筛选的列筛选条件。|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks () ](/javascript/api/excel/excel.linkedworkbook#breakLinks__)|请求断开指向链接工作簿的链接。|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|指向链接工作簿的原始 URL。|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh__)|请求刷新从链接工作簿检索到的数据。|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks () ](/javascript/api/excel/excel.linkedworkbookcollection#breakAllLinks__)|断开指向链接工作簿的所有链接。|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItem_key_)|按 URL 获取链接工作簿的信息。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItemOrNullObject_key_)|按 URL 获取链接工作簿的信息。|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|获取此集合中已加载的子项。|
||[refreshAll () ](/javascript/api/excel/excel.linkedworkbookcollection#refreshAll__)|请求刷新所有工作簿链接。|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbookLinksRefreshMode)|表示工作簿链接的更新模式。|
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
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedWorkbooks)|返回链接工作簿的集合。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedSheetViews)|返回工作表中呈现的工作表视图的集合。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Excel JavaScript 预览 API](excel-preview-apis.md)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)

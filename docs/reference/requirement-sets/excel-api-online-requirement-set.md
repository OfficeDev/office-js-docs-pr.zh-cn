---
title: ExcelJavaScript API 仅联机要求集
description: 有关 ExcelApiOnline 要求集的详细信息。
ms.date: 10/14/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: bd1f62b92b9a08d23daf77f8f4b86c60333faab3
ms.sourcegitcommit: e4d98eb90e516b9c90e3832f3212caf48691acf6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/22/2021
ms.locfileid: "60537588"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>ExcelJavaScript API 仅联机要求集

要求 `ExcelApiOnline` 集是一个特殊要求集，其中包括仅适用于Excel web 版。 此要求集内 API 被视为生产 API， (应用程序未记录的行为或) 更改Excel web 版 API。 `ExcelApiOnline`API 被视为适用于其他平台（如 (Windows、Mac、iOS) ）的"预览"API，可能不受这些平台的任何支持。

当所有平台都支持要求集内 API 时，它们将被添加到下一个发布的要求 `ExcelApiOnline` `ExcelApi 1.[NEXT]` () 。 一旦该新要求公开，将从 中删除这些 `ExcelApiOnline` API。 将此过程视为与从预览版移动到发布的 API 类似的推广过程。

> [!IMPORTANT]
> `ExcelApiOnline` 是最新编号要求集的超集。

> [!IMPORTANT]
> `ExcelApiOnline 1.1` 是仅联机 API 的唯一版本。 这是因为Excel web 版将始终为最新版本的用户提供单个版本。

下表提供了 API 的简要摘要，而后续 API 列表表提供了当前 [API](#api-list) `ExcelApiOnline` 的详细列表。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 链接的工作簿 | 管理工作簿之间的链接，包括对刷新和断开工作簿链接的支持。 | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook) [、LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| 命名工作表视图 | 以编程方式控制每用户工作表视图。 | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) [、NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection) |
| 工作表移动事件 | 检测何时在集合中移动工作表、工作表的位置以及更改的来源。 | [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) [、WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs) |

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
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows (行： number[] \| TableRow[]) ](/javascript/api/excel/excel.tablerowcollection#deleteRows_rows_)|从表中删除多行。|
||[deleteRowsAt (索引： number， count？： number) ](/javascript/api/excel/excel.tablerowcollection#deleteRowsAt_index__count_)|从给定索引开始，从表中删除指定数量的行。|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedWorkbooks)|返回链接工作簿的集合。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedSheetViews)|返回工作表中呈现的工作表视图的集合。|
||[onNameChanged](/javascript/api/excel/excel.worksheet#onNameChanged)|更改工作表名称时发生。|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#onVisibilityChanged)|在工作表可见性更改时发生。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onMoved](/javascript/api/excel/excel.worksheetcollection#onMoved)|当用户在工作簿中移动工作表时发生。|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#onNameChanged)|在工作表集合中更改工作表名称时发生。|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#onVisibilityChanged)|在工作表集合中更改工作表可见性时发生。|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#positionAfter)|在移动后获取工作表的新位置。|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#positionBefore)|在移动之前获取工作表的上一位置。|
||[源](/javascript/api/excel/excel.worksheetmovedeventargs#source)|事件的源。|
||[type](/javascript/api/excel/excel.worksheetmovedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#worksheetId)|获取已移动工作表的 ID。|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameAfter)|在名称更改后获取工作表的新名称。|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameBefore)|获取工作表的先前名称，在名称更改之前。|
||[源](/javascript/api/excel/excel.worksheetnamechangedeventargs#source)|事件的源。|
||[type](/javascript/api/excel/excel.worksheetnamechangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#worksheetId)|获取具有新名称的工作表的 ID。|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[源](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#source)|事件的源。|
||[type](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#type)|获取事件的类型。|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityAfter)|获取在可见性更改后工作表的新可见性设置。|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityBefore)|获取工作表的上一个可见性设置，在可见性更改之前。|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#worksheetId)|获取其可见性已更改的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Excel JavaScript 预览 API](excel-preview-apis.md)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)

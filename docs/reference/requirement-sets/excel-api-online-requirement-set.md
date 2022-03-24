---
title: Excel JavaScript API 仅联机要求集
description: 有关 ExcelApiOnline 要求集的详细信息。
ms.date: 10/29/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: f3ec510e889ecfe565767352c59cd349e0701830
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746603"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel JavaScript API 仅联机要求集

要求`ExcelApiOnline`集是一个特殊要求集，其中包括仅适用于Excel web 版。 此要求集内 API 被视为生产 API， (应用程序未记录的行为或) 更改Excel web 版 API。 `ExcelApiOnline`API 被视为其他平台（如 (Windows、Mac、iOS) ）的"预览"API，可能不受这些平台的任何支持。

当要求集内 `ExcelApiOnline` `ExcelApi 1.[NEXT]` API 在所有平台中均受支持时，它们将被添加到下一个发布的要求 () 。 一旦该新要求公开，将从 中删除这些 API `ExcelApiOnline`。 将此过程视为与从预览版移动到发布的 API 类似的推广过程。

> [!IMPORTANT]
> `ExcelApiOnline` 是最新编号要求集的超集。

> [!IMPORTANT]
> `ExcelApiOnline 1.1` 是仅联机 API 的唯一版本。 这是因为Excel web 版将始终为最新版本的用户提供单个版本。

下表提供了 API 的简要摘要，而后续 API 列表表提供了当前 [API](#api-list) 的详细 `ExcelApiOnline` 列表。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 链接的工作簿 | 管理工作簿之间的链接，包括对刷新和断开工作簿链接的支持。 | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)、 [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| 命名工作表视图 | 以编程方式控制每用户工作表视图。 | [NamedSheetView](/javascript/api/excel/excel.namedsheetview)、 [NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection) |
| 工作表移动事件 | 检测何时在集合中移动工作表、工作表的位置以及更改的来源。 | [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)、 [WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs) |

## <a name="recommended-usage"></a>建议的用法

由于 `ExcelApiOnline` API 仅受 Excel web 版 支持，因此加载项应在调用这些 API 之前检查要求集是否受支持。 这可以避免在不同的平台上调用仅联机 API。

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

API 位于跨平台要求集后，应删除或编辑 `isSetSupported` 检查。 这将在其他平台上启用外接程序的功能。 进行此更改时，请务必在这些平台上测试该功能。

> [!IMPORTANT]
> 清单不能指定为 `ExcelApiOnline 1.1` 激活要求。 它在 Set 元素中不是有效的 [值](../manifest/set.md)。

## <a name="api-list"></a>API 列表

下表列出了要求Excel中当前包含的 `ExcelApiOnline` JavaScript API。 有关所有 JavaScript `ExcelApiOnline` EXCEL的完整列表 (包括 API 和以前发布的 API) ，请参阅所有 Excel [JavaScript API](/javascript/api/excel?view=excel-js-online&preserve-view=true)。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks () ](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-breaklinks-member(1))|请求断开指向链接工作簿的链接。|
||[id](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-id-member)|指向链接工作簿的原始 URL。|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-refresh-member(1))|请求刷新从链接工作簿检索到的数据。|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks () ](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-breakalllinks-member(1))|断开指向链接工作簿的所有链接。|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-getitem-member(1))|按 URL 获取链接工作簿的信息。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-getitemornullobject-member(1))|按 URL 获取链接工作簿的信息。|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-items-member)|获取此集合中已加载的子项。|
||[refreshAll () ](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-refreshall-member(1))|请求刷新所有工作簿链接。|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-workbooklinksrefreshmode-member)|表示工作簿链接的更新模式。|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-activate-member(1))|激活此工作表视图。|
||[delete()](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-delete-member(1))|从工作表中删除工作表视图。|
||[duplicate (name？： string) ](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-duplicate-member(1))|创建此工作表视图的副本。|
||[name](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-name-member)|获取或设置工作表视图的名称。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-add-member(1))|创建具有给定名称的新工作表视图。|
||[enterTemporary () ](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-entertemporary-member(1))|创建并激活新的临时工作表视图。|
||[exit () ](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-exit-member(1))|退出当前活动的工作表视图。|
||[getActive () ](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getactive-member(1))|获取工作表当前的活动工作表视图。|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getcount-member(1))|获取此工作表中的工作表视图数。|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitem-member(1))|使用工作表视图的名称获取工作表视图。|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitemat-member(1))|按工作表视图在集合中的索引获取工作表视图。|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-items-member)|获取此集合中已加载的子项。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows (行： number[] \| TableRow[]) ](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-deleterows-member(1))|从表中删除多行。|
||[deleteRowsAt (索引： number， count？： number) ](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-deleterowsat-member(1))|从给定索引开始，从表中删除指定数量的行。|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkedworkbooks-member)|返回链接工作簿的集合。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-namedsheetviews-member)|返回工作表中呈现的工作表视图的集合。|
||[onNameChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onnamechanged-member)|更改工作表名称时发生。|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onvisibilitychanged-member)|在工作表可见性更改时发生。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onMoved](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onmoved-member)|当用户在工作簿中移动工作表时发生。|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onnamechanged-member)|在工作表集合中更改工作表名称时发生。|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onvisibilitychanged-member)|在工作表集合中更改工作表可见性时发生。|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-positionafter-member)|在移动后获取工作表的新位置。|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-positionbefore-member)|在移动之前获取工作表的上一位置。|
||[源](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-source-member)|事件的源。|
||[type](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-worksheetid-member)|获取已移动工作表的 ID。|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-nameafter-member)|在名称更改后获取工作表的新名称。|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-namebefore-member)|获取工作表的先前名称，在名称更改之前。|
||[源](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-source-member)|事件的源。|
||[type](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-worksheetid-member)|获取具有新名称的工作表的 ID。|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[源](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-source-member)|事件的源。|
||[type](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-type-member)|获取事件的类型。|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-visibilityafter-member)|获取在可见性更改后工作表的新可见性设置。|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-visibilitybefore-member)|获取工作表的上一个可见性设置，在可见性更改之前。|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-worksheetid-member)|获取其可见性已更改的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Excel JavaScript 预览 API](excel-preview-apis.md)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)

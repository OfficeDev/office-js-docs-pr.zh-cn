---
title: 加载项Excel疑难解答
description: 了解如何解决加载项中的Excel错误。
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: c6a523354cc938ac9e9ba041ddb09f12142a3a58
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340790"
---
# <a name="troubleshooting-excel-add-ins"></a>加载项Excel疑难解答

本文讨论对解决方案唯一的Excel。 请使用页面底部的反馈工具，建议可添加到文章中的其他问题。

## <a name="api-limitations-when-the-active-workbook-switches"></a>活动工作簿切换时的 API 限制

加载项Excel一次对一个工作簿进行操作。 与运行加载项的工作簿分开的工作簿获得焦点时，可能会出现错误。 只有在焦点更改时调用特定方法时，才会发生此情况。

以下 API 受此工作簿开关的影响。

|Excel JavaScript API | 抛出的错误 |
|--|--|
| `Chart.activate` | GeneralException |
| `Range.select` | GeneralException |
| `Table.clearFilters` | GeneralException |
| `Workbook.getActiveCell`  | InvalidSelection|
| `Workbook.getSelectedRange` | InvalidSelection|
| `Workbook.getSelectedRanges`  | InvalidSelection|
| `Worksheet.activate` | GeneralException |
| `Worksheet.delete`  | InvalidSelection|
| `Worksheet.gridlines` | GeneralException |
| `Worksheet.showHeadings` | GeneralException |
| `WorksheetCollection.add` | GeneralException |
| `WorksheetFreezePanes.freezeAt` | GeneralException |
| `WorksheetFreezePanes.freezeColumns` | GeneralException |
| `WorksheetFreezePanes.freezeRows` | GeneralException |
| `WorksheetFreezePanes.getLocationOrNullObject`| GeneralException |
| `WorksheetFreezePanes.unfreeze` | GeneralException |

> [!NOTE]
> 这仅适用于在 Excel Mac 上打开的多个Windows工作簿。

## <a name="coauthoring"></a>共同创作

请参阅[共同Excel](co-authoring-in-excel-add-ins.md)外接程序中的共同授权，了解用于共同授权环境中事件的模式。 本文还讨论了使用某些 API（如 ）时的潜在合并冲突 [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1))。

## <a name="known-issues"></a>已知问题

### <a name="binding-events-return-temporary-binding-obects"></a>绑定事件返回临时 `Binding` 对象

[BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#excel-excel-bindingdatachangedeventargs-binding-member) 和 [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#excel-excel-bindingselectionchangedeventargs-binding-member) `Binding` 都返回一个临时对象，其中包含引发该事件的对象的 ID`Binding`。 使用此 ID 检索 `BindingCollection.getItem(id)` 引发 `Binding` 事件的对象。

下面的代码示例演示如何使用此临时绑定 ID 检索相关 `Binding` 对象。 在示例中，将事件侦听器分配给绑定。 侦听器在触发 `getBindingId` 事件 `onDataChanged` 时调用 方法。 方法 `getBindingId` 使用临时对象的 `Binding` ID 检索引发 `Binding` 事件的对象。

```js
async function run() {
    await Excel.run(async (context) => {
        // Retrieve your binding.
        let binding = context.workbook.bindings.getItemAt(0);
    
        await context.sync();
    
        // Register an event listener to detect changes to your binding
        // and then trigger the `getBindingId` method when the data changes. 
        binding.onDataChanged.add(getBindingId);
        await context.sync();
    });
}

async function getBindingId(eventArgs) {
    await Excel.run(async (context) => {
        // Get the temporary binding object and load its ID. 
        let tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        let originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // You now have the binding object that raised the event: `originalBindingObject`. 
    });
}
```

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a>单元格格式 `useStandardHeight` 和 `useStandardWidth` 问题

[的 useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member) `CellPropertiesFormat` 属性在属性中Excel web 版。 由于用户界面中Excel web 版问题`useStandardHeight``true`，因此将 属性设置为不精确地在此平台上计算高度。 例如，标准高度 **14** 在 Excel web 版 中修改为 **14.25**。

在所有平台上， [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member) 和 [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardwidth-member) `CellPropertiesFormat` 属性仅旨在设置为 `true`。 将这些属性设置为 `false` 无效。

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a>区域`getImage`方法不受支持Excel for Mac

Range [getImage](/javascript/api/excel/excel.range#excel-excel-range-getimage-member(1)) 方法当前在 Excel for Mac。 有关 [当前状态，请参阅 OfficeDev/office-js Issue #235](https://github.com/OfficeDev/office-js/issues/235) 。

### <a name="range-return-character-limit"></a>区域返回字符限制

[Worksheet.getRange (address) ](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1)) [和 Worksheet.getRanges (address) ](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getranges-member(1)) 方法的地址字符串限制为 8192 个字符。 超出此限制时，地址字符串将被截断为 8192 个字符。

## <a name="see-also"></a>另请参阅

- [排查 Office 加载项中的开发错误](../testing/troubleshoot-development-errors.md)
- [排查 Office 加载项中的用户错误](../testing/testing-and-troubleshooting.md)

---
title: Excel 加载项疑难解答
description: 了解如何解决 Excel 加载项中的开发错误。
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: 0efc8b4d25d9d748975146e187104972e4ad58a9
ms.sourcegitcommit: 1cdf5728102424a46998e1527508b4e7f9f74a4c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50270726"
---
# <a name="troubleshooting-excel-add-ins"></a>Excel 加载项疑难解答

本文讨论 Excel 特有的疑难解答问题。 请使用页面底部的反馈工具建议可添加到文章中的其他问题。

## <a name="api-limitations-when-the-active-workbook-switches"></a>活动工作簿切换时的 API 限制

Excel 外接程序旨在一次对一个工作簿进行操作。 当与运行加载项的工作簿分开的工作簿获得焦点时，可能会出现错误。 只有在焦点更改时调用特定方法时，才会发生此情况。

以下 API 受此工作簿开关的影响：

|Excel JavaScript API | 引发错误 |
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
> 这仅适用于在 Windows 或 Mac 上打开的多个 Excel 工作簿。

## <a name="coauthoring"></a>共同创作

请参阅 [Excel 加载项中的](co-authoring-in-excel-add-ins.md) 共同授权，了解用于共同授权环境中事件的模式。 本文还讨论使用某些 API 时的潜在合并冲突，例如 [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) 。

## <a name="known-issues"></a>已知问题

### <a name="binding-events-return-temporary-binding-obects"></a>绑定事件返回临时 `Binding` 对象

[BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding)和[BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding)都返回一个临时对象，该对象包含引发该事件 `Binding` 的对象的 `Binding` ID。 使用此 ID `BindingCollection.getItem(id)` 检索 `Binding` 引发事件的对象。

下面的代码示例演示如何使用此临时绑定 ID 检索相关 `Binding` 对象。 在示例中，将事件侦听器分配给绑定。 当触发 `getBindingId` 事件时，侦听器 `onDataChanged` 将调用该方法。 `getBindingId`该方法使用临时对象的 ID 检索 `Binding` `Binding` 引发事件的对象。

```js
Excel.run(function (context) {
    // Retrieve your binding.
    var binding = context.workbook.bindings.getItemAt(0);

    return context.sync().then(function () {
        // Register an event listener to detect changes to your binding
        // and then trigger the `getBindingId` method when the data changes. 
        binding.onDataChanged.add(getBindingId);

        return context.sync();
    });
});

function getBindingId(eventArgs) {
    return Excel.run(function (context) {
        // Get the temporary binding object and load its ID. 
        var tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        var originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // You now have the binding object that raised the event: `originalBindingObject`. 
    });
}
```

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a>单元格格式 `useStandardHeight` `useStandardWidth` 和问题

[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight)属性在 `CellPropertiesFormat` Excel 网页中无法正常工作。 由于 Excel 网页 UI 中的问题，将该属性设置为在此平台上计算高度不 `useStandardHeight` `true` 精确。 例如，在 Excel 网页版中，标准高度 **14** 修改为 **14.25。**

在所有平台上 [，useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) 和 [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) 属性仅用于 `CellPropertiesFormat` 设置为 `true` 。 将这些属性设置为 `false` 不起作用。 

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a>Excel `getImage` for Mac 不支持 Range 方法

Excel for Mac 当前不支持 Range [getImage](/javascript/api/excel/excel.range#getImage__) 方法。 请参阅 [OfficeDev/office-js #235](https://github.com/OfficeDev/office-js/issues/235) 当前状态。

### <a name="range-return-character-limit"></a>区域返回字符限制

[Worksheet.getRange (address) ](/javascript/api/excel/excel.worksheet#getRange_address_) [和 Worksheet.getRanges (address) ](/javascript/api/excel/excel.worksheet#getRanges_address_)方法的地址字符串限制为 8192 个字符。 超过此限制时，地址字符串将被截断为 8192 个字符。

## <a name="see-also"></a>另请参阅

- [Office 加载项的开发错误疑难解答](../testing/troubleshoot-development-errors.md)
- [排查 Office 加载项中的用户错误](../testing/testing-and-troubleshooting.md)

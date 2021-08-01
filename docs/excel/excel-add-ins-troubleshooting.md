---
title: 加载项Excel疑难解答
description: 了解如何解决加载项中的Excel错误。
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: b90d8cfdb4696445655122a2fa7eb74d1c87fa2f
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671462"
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

有关[用于共同Excel](co-authoring-in-excel-add-ins.md)中的事件的模式，请参阅在加载项中共同授权。 本文还讨论了使用某些 API（如 ）时的潜在合并冲突 [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add_index__values_) 。

## <a name="known-issues"></a>已知问题

### <a name="binding-events-return-temporary-binding-obects"></a>绑定事件返回 `Binding` 临时对象

[BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding)和[BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding)都返回一个临时对象，其中包含引发 `Binding` `Binding` 该事件的对象的 ID。 使用此 ID 检索 `BindingCollection.getItem(id)` `Binding` 引发事件的对象。

下面的代码示例演示如何使用此临时绑定 ID 检索相关 `Binding` 对象。 在示例中，将事件侦听器分配给绑定。 侦听器在 `getBindingId` 触发事件 `onDataChanged` 时调用 方法。 `getBindingId`方法使用临时对象的 ID `Binding` 检索 `Binding` 引发事件的对象。

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

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a>单元格格式 `useStandardHeight` 和 `useStandardWidth` 问题

[的 useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight)属性在属性中 `CellPropertiesFormat` Excel web 版。 由于用户界面中Excel web 版问题，因此将 属性设置为不精确地在此平台上 `useStandardHeight` `true` 计算高度。 例如，标准高度 **14** 在 Excel web 版 中修改为 **14.25。**

在所有平台上 [，useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) 和 [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) 属性仅旨在 `CellPropertiesFormat` 设置为 `true` 。 将这些属性设置为 `false` 无效。 

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a>区域 `getImage` 方法不受支持Excel for Mac

Range [getImage](/javascript/api/excel/excel.range#getImage__)方法当前在 Excel for Mac。 请参阅 [OfficeDev/office-js issue #235](https://github.com/OfficeDev/office-js/issues/235) 了解当前状态。

### <a name="range-return-character-limit"></a>区域返回字符限制

[Worksheet.getRange (address) ](/javascript/api/excel/excel.worksheet#getRange_address_) [和 Worksheet.getRanges](/javascript/api/excel/excel.worksheet#getRanges_address_) (address) 方法的地址字符串限制为 8192 个字符。 超出此限制时，地址字符串将被截断为 8192 个字符。

## <a name="see-also"></a>另请参阅

- [排查Office加载项的开发错误](../testing/troubleshoot-development-errors.md)
- [排查 Office 加载项中的用户错误](../testing/testing-and-troubleshooting.md)

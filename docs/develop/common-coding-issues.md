---
title: 常见问题和意外平台行为的编码指南
description: 开发人员经常遇到的 Office JavaScript API 平台问题的列表。
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: f6d6a31059b32550e3176ed278d7da4c2c7a6c68
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292909"
---
# <a name="coding-guidance-for-common-issues-and-unexpected-platform-behaviors"></a>常见问题和意外平台行为的编码指南

本文重点介绍了 Office JavaScript API 的各个方面，这些方面可能导致意外行为或需要特定编码模式来实现所需的结果。 如果遇到此列表中的问题，请使用文章底部的反馈表单告知我们。

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a>通用 Api 和 Outlook Api 不基于承诺

[公共 api](/javascript/api/office) (那些与特定 Office 应用程序不关联的 api) 并且[Outlook api](/javascript/api/outlook)使用基于回调的编程模型。 与基础 Office 文档进行交互需要进行异步读取或写入调用，以指定在操作完成时运行的回调。 有关此模式的示例，请参阅 [document.getfileasync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)。

这些常见 API 和 Outlook API 方法不会返回 [承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)。 因此，在异步操作完成之前，不能使用 [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) 暂停执行。 如果需要 `await` 行为，可以在显式创建的承诺中包装方法调用。

```js
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

> [!NOTE]
> 参考文档包含 [getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-)的承诺包装实现。

## <a name="some-properties-cannot-be-set-directly"></a>某些属性不能直接设置

> [!NOTE]
> 本部分仅适用于适用于 Excel 和 Word 的应用程序特定的 Api。

某些属性虽然是可写的，但不能设置。 这些属性是父属性的一部分，必须将其设置为单个对象。 这是因为该父属性依赖具有特定逻辑关系的子属性。 必须使用对象文本表示法设置这些父属性，以设置整个对象，而不是设置该对象的单个子属性。 在 [页面布局](/javascript/api/excel/excel.pagelayout)中找到此示例的一个示例。 `zoom`必须使用单个[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)对象设置属性，如下所示：

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

在上面的示例中，您 ***将无法*** 直接分配 `zoom` 值： `sheet.pageLayout.zoom.scale = 200;` 。 由于未加载，该语句 `zoom` 会引发错误。 即使 `zoom` 要加载，该扩展集也不会生效。 发生所有上下文操作 `zoom` ，刷新加载项中的代理对象并覆盖本地设置的值。

此行为不同于 [导航属性](application-specific-api-model.md#scalar-and-navigation-properties) ，如 [Range. 格式](/javascript/api/excel/excel.range#format)。 `format`可以使用对象导航设置属性，如下所示：

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

您可以通过检查属性的只读修饰符来标识无法直接设置其子属性的属性。 所有只读属性都可以直接设置其非只读的子属性。 `PageLayout.zoom`必须使用该级别的对象设置可写属性（如必须设置）。 摘要：

- 只读属性：可通过导航设置子属性。
- 可写属性：不能通过导航 (设置子属性，而必须将初始父对象分配) 的一部分。

## <a name="setting-read-only-properties"></a>设置只读属性

Office JS 的 [TypeScript 定义](referencing-the-javascript-api-for-office-library-from-its-cdn.md) 指定哪些对象属性是只读的。 如果尝试设置只读属性，写入操作将无提示地失败，且不会引发错误。 下面的示例错误地尝试设置只读属性 [Chart.id](/javascript/api/excel/excel.chart#id)。

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="removing-event-handlers"></a>删除事件处理程序

必须使用在其中添加事件处理程序的相同项将其删除 `RequestContext` 。 如果需要加载项在运行时删除事件处理程序，则需要存储用于添加处理程序的 context 对象。

```js
Excel.run(async (context) => {
    [...]

    // To later remove an event handler, store the context somewhere accessible to the handler removal function.
    // You may find it helpful to also store the event handler object and associate it with the context.
    selectionChangedHandler = myWorksheet.onSelectionChanged.add(callback);
    savedContext = currentContext;
    return context.sync();
}
```

## <a name="supporting-internet-explorer"></a>支持 Internet Explorer

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="excel-specific-issues"></a>特定于 Excel 的问题

### <a name="api-limitations-when-the-active-workbook-switches"></a>活动工作簿切换时的 API 限制

Excel 相关外接程序用于一次运行单个工作簿。 当运行加载项的工作簿获得焦点时，可能会出现错误。 仅当焦点更改时要调用的特定方法时，才会发生这种情况。

此工作簿开关会影响以下 Api：

|Excel JavaScript API | 引发的错误 |
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

### <a name="coauthoring"></a>共同创作

请参阅 [Excel 外接程序中](../excel/co-authoring-in-excel-add-ins.md) 用于共同创作环境中事件的模式的合著。 本文还讨论了使用某些 Api （例如）时的潜在合并冲突 [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) 。

## <a name="see-also"></a>另请参阅

- [Office 外接程序的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md)
- [OfficeDev/？ js](https://github.com/OfficeDev/office-js/issues)：报告和查看 office 外接程序平台和 JavaScript api 中的问题的位置。
- [堆栈溢出](https://stackoverflow.com/questions/tagged/office-js)：询问并查看有关 Office JavaScript api 的编程问题的位置。 在发布到堆栈溢出时，请务必对您的问题应用 "office-js" 标记。
- [UserVoice](https://officespdev.uservoice.com/)：建议 Office 外接程序平台和 Office JavaScript api 的新功能的位置。

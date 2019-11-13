---
title: 常见的编码问题和意外的平台行为
description: 开发人员经常遇到的 Office JavaScript API 平台问题的列表。
ms.date: 11/06/2019
localization_priority: Normal
ms.openlocfilehash: a4d7a09c1645bea181060157d933036d1924044f
ms.sourcegitcommit: 88d81aa2d707105cf0eb55d9774b2e7cf468b03a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/13/2019
ms.locfileid: "38301930"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a>常见的编码问题和意外的平台行为

本文重点介绍了 Office JavaScript API 的各个方面，这些方面可能导致意外行为或需要特定编码模式来实现所需的结果。 如果遇到此列表中的问题，请使用文章底部的反馈表单告知我们。

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a>通用 Api 和 Outlook Api 不基于承诺

[通用 api](/javascript/api/office) （那些未绑定到特定 Office 主机的 api）和[Outlook api](/javascript/api/outlook)使用基于回调的编程模型。 与基础 Office 文档进行交互需要进行异步读取或写入调用，以指定在操作完成时要运行的回调。 有关此模式的示例，请参阅[document.getfileasync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)。

这些常见 API 和 Outlook API 方法不会返回[承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)。 因此，在异步操作完成之前，不能使用[await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await)暂停执行。 如果需要`await`行为，可以在显式创建的承诺中包装方法调用。

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
> 参考文档包含[getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-)的承诺包装实现。

## <a name="some-properties-must-be-set-with-json-structs"></a>某些属性必须使用 JSON 结构进行设置

> [!NOTE]
> 本部分仅适用于 Excel 和 Word 的特定于主机的 Api。

某些属性必须设置为 JSON 结构，而不是设置其单个子属性。 在[页面布局](/javascript/api/excel/excel.pagelayout)中找到此示例的一个示例。 必须`zoom`使用单个[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)对象设置属性，如下所示：

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

在上面的示例中，您***将无法***直接分配`zoom`值： `sheet.pageLayout.zoom.scale = 200;`。 由于`zoom`未加载，该语句会引发错误。 `zoom`即使要加载，该扩展集也不会生效。 发生所有上下文操作`zoom`，刷新加载项中的代理对象并覆盖本地设置的值。

此行为不同于[导航属性](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties)，如[Range. 格式](/javascript/api/excel/excel.range#format)。 `format`可以使用对象导航设置属性，如下所示：

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

您可以通过检查其只读修饰符来标识必须将其子属性设置为 JSON 结构的属性。 所有只读属性都可以直接设置其非只读的子属性。 必须使用 JSON `PageLayout.zoom`结构设置可写属性（如必须设置）。 摘要：

- 只读属性：可通过导航设置子属性。
- 可写属性：必须使用 JSON 结构设置子属性（且不能通过导航进行设置）。

## <a name="excel-data-transfer-limits"></a>Excel 数据传输限制

如果您正在构建 Excel 外接程序，请注意与工作簿交互时的以下大小限制：

- Excel 网页版将请求和响应的有效负载大小限制为 5MB。 如果超过该限制，将引发 `RichAPI.Error`。
- 对于 get 操作，范围限制为5000000个单元格。

如果您希望用户输入超出这些限制，请务必先检查数据，然后再调用`context.sync()`。 根据需要将操作拆分为较小的部分。 请务必为每`context.sync()`个子操作调用，以避免这些操作再次成批组合。

这些限制通常由大型区域所超过。 您的外接程序可能能够使用[RangeAreas](/javascript/api/excel/excel.rangeareas)对较大范围内的单元格进行战略更新。 有关详细信息，请参阅[在 Excel 外接程序中同时处理多个区域](../excel/excel-add-ins-multiple-ranges.md)。

## <a name="setting-read-only-properties"></a>设置只读属性

Office JS 的[TypeScript 定义](/referencing-the-javascript-api-for-office-library-from-its-cdn.md)指定哪些对象属性是只读的。 如果尝试设置只读属性，写入操作将无提示地失败，且不会引发错误。 下面的示例错误地尝试设置只读属性[Chart.id](/javascript/api/excel/excel.chart#id)。

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a>另请参阅

- [OfficeDev/？ js](https://github.com/OfficeDev/office-js/issues)：报告和查看 office 外接程序平台和 JavaScript api 中的问题的位置。
- [堆栈溢出](https://stackoverflow.com/questions/tagged/office-js)：询问并查看有关 Office JavaScript api 的编程问题的位置。 在发布到堆栈溢出时，请务必对您的问题应用 "office-js" 标记。
- [UserVoice](https://officespdev.uservoice.com/)：建议 Office 外接程序平台和 Office JavaScript api 的新功能的位置。

---
title: 使用 Excel JavaScript API 处理错误
description: 了解Excel用于解释运行时错误的 JavaScript API 错误处理逻辑。
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6fa5ca0c7ebf9400fcdd83c7bf4eb4b906f2e5b5
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/15/2022
ms.locfileid: "66090829"
---
# <a name="error-handling-with-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理错误

使用 Excel JavaScript API 生成加载项时，请务必加入错误处理逻辑，以便解决运行时错误。 鉴于 API 的异步特性，这样做非常关键。

> [!NOTE]
> 有关 `sync()` Excel JavaScript API 的方法和异步特性的详细信息，请参阅 [Office 外接程序中的 Excel JavaScript 对象模型](excel-add-ins-core-concepts.md)。

## <a name="best-practices"></a>最佳做法

在我们的[代码示例](https://github.com/OfficeDev/Office-Add-in-samples)和[Script Lab](../overview/explore-with-script-lab.md)代码片段中，你会注意到每次调用`Excel.run`都附带一个`catch`语句来捕获在其中`Excel.run`发生的任何错误。 建议在使用 Excel JavaScript API 生成加载项时使用相同模式。

```js
$("#run").click(() => tryCatch(run));

async function run() {
  await Excel.run(async (context) => {
      // Add your Excel JavaScript API calls here.

      // Await the completion of context.sync() before continuing.
    await context.sync();
    console.log("Finished!");
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

```

## <a name="api-errors"></a>API 错误

当 Excel JavaScript API 请求无法成功运行时，API 将返回包含以下属性的错误对象。

- **代码**：错误消息的 `code` 属性包含一个字符串，它属于 `OfficeExtension.ErrorCodes` 或 `Excel.ErrorCodes` 列表的一部分。 例如，错误代码“InvalidReference”指示引用对于指定操作无效。 错误代码尚未本地化。

- **消息**：错误消息的 `message` 属性包含本地化字符串中的错误摘要。 错误消息并非针对最终用户的使用情况；应使用错误代码和相应的业务逻辑来确定外接程序显示给最终用户的错误消息。

- **debugInfo**：出现此信息时，错误消息的 `debugInfo` 属性将提供其他信息，帮助理解错误根本原因。

> [!NOTE]
> 如果用于 `console.log()` 将错误消息打印到控制台，则这些消息仅在服务器上可见。 最终用户不会在加载项任务窗格或Office应用程序中的任何位置看到这些错误消息。 若要向用户报告错误，请参阅 [错误通知](#error-notifications)。

## <a name="error-messages"></a>错误消息

下表是 API 可能返回的错误列表。

|错误代码 | 错误消息 | 注释 |
|:----------|:--------------|:------|
|`AccessDenied` |无法执行所请求的操作。| |
|`ActivityLimitReached`|已达到活动限制。| |
|`ApiNotAvailable`|请求的 API 不可用。| |
|`ApiNotFound`|找不到你尝试使用的 API。 它可能在较新版本的Excel中提供。 有关详细信息，请参阅 [Excel JavaScript API 要求集](/javascript/api/requirement-sets/excel/excel-api-requirement-sets)文章。| |
|`BadPassword`|提供的密码不正确。| |
|`Conflict`|由于冲突，无法处理请求。| |
|`ContentLengthRequired`|`Content-length`缺少 HTTP 标头。| |
|`EmptyChartSeries`|尝试的操作失败，因为图表系列为空。| |
|`FilteredRangeConflict`|尝试的操作会导致与筛选范围冲突。| |
|`FormulaLengthExceedsLimit`|应用公式的字节码超出了最大长度限制。 对于 32 位计算机上的Office，字节码长度限制为 16384 个字符。 在 64 位计算机上，字节码长度限制为 32768 个字符。| 此错误发生在Excel web 版和桌面上。|
|`GeneralException`|处理请求时出现内部错误。| |
|`InactiveWorkbook`|操作失败，因为多个工作簿处于打开状态，并且此 API 调用的工作簿已失去焦点。| |
|`InsertDeleteConflict`|尝试的插入或删除操作导致冲突。| |
|`InvalidArgument` |自变量无效、缺少或格式不正确。| |
|`InvalidBinding` |由于之前的更新，此对象绑定不再有效。| |
|`InvalidOperation`|尝试的操作对于对象无效。| |
|`InvalidOperationInCellEditMode`|当Excel处于“编辑单元格”模式时，该操作不可用。 使用 **Enter** 或 **Tab** 键或选择其他单元格退出编辑模式，然后重试。| |
|`InvalidReference`|此引用对于当前操作无效。| |
|`InvalidRequest`  |无法处理此请求。| |
|`InvalidSelection`|当前选定内容对于此操作无效。| |
|`ItemAlreadyExists`|所创建的资源已存在。| |
|`ItemNotFound` |所请求的资源不存在。| |
|`MemoryLimitReached`|已达到内存限制。 无法完成操作。| |
|`MergedRangeConflict`|无法完成操作。 表不能与其他表、数据透视表、查询结果、合并单元格或 XML 映射重叠。|
|`NonBlankCellOffSheet`|Microsoft Excel无法插入新单元格，因为它会将非空单元格推离工作表的末尾。 这些非空单元格可能显示为空，但具有空白值、某些格式或公式。 删除足够的行或列，为要插入的内容腾出空间，然后重试。| |
|`NotImplemented`|所请求的功能未实现。| |
|`OperationCellsExceedLimit`|尝试的操作影响超过 33554000 个单元格的限制。| `TableColumnCollection.add API`如果触发此错误，请确认工作表中没有意外数据，但在表外部。 特别是，检查工作表最右侧列中的数据。 删除意外数据以解决此错误。 验证操作进程的单元格数的一种方法是运行以下计算： `(number of table rows) x (16383 - (number of table columns))` 数字 16383 是Excel支持的最大列数。 <br><br>此错误仅在Excel web 版中发生。 |
|`PivotTableRangeConflict`|尝试的操作会导致与数据透视表范围冲突。| |
|`RangeExceedsLimit`|范围内的单元格计数已超过支持的最大数目。 有关详细信息，请参阅[Office加载项文章的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)。| |
|`RefreshWorkbookLinksBlocked`|操作失败，因为用户尚未授予刷新外部工作簿链接的权限。| |
|`RequestAborted`|请求在运行时已中止。| |
|`RequestPayloadSizeLimitExceeded`|请求有效负载大小已超出限制。 有关详细信息，请参阅[Office加载项文章的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)。| 此错误仅在Excel web 版中发生。|
|`ResponsePayloadSizeLimitExceeded`|响应有效负载大小已超出限制。 有关详细信息，请参阅[Office加载项文章的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)。|  此错误仅在Excel web 版中发生。|
|`ServiceNotAvailable`|服务不可用。| |
|`Unauthenticated` |所需的身份验证信息缺少或无效。| |
|`UnsupportedFeature`|操作失败，因为源工作表包含一个或多个不受支持的功能。| |
|`UnsupportedOperation`|不支持正在尝试的操作。| |
|`UnsupportedSheet`|此工作表类型不支持此操作，因为它是宏或图表工作表。| |

> [!NOTE]
> 上表列出了使用 Excel JavaScript API 时可能会遇到的错误消息。 如果使用 Common API 而不是特定于应用程序的 Excel JavaScript API，请参阅[Office常见 API 错误代码](../reference/javascript-api-for-office-error-codes.md)，了解相关错误消息。

## <a name="error-notifications"></a>错误通知

如何向用户报告错误取决于所使用的 UI 系统。 如果使用React作为 UI 系统，请使用Fluent UI 组件和设计元素。 从此[Fluent UI 页面](https://developer.microsoft.com/fluentui#/controls/web)中选择适当的控件。 建议使用消息栏、对话框或模式传达错误消息。 如果错误位于用户输入中，则在输入控件附近以粗红色显示错误。 示例[Office加载项-Microsoft-Graph-React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/0706cc67645675a48747f1fec1b1e5722b575b11/Samples/auth/Office-Add-in-Microsoft-Graph-React)使用 MessageBar 元素并对其进行修改，以考虑加载项任务窗格中的个性菜单。

如果不对 UI 使用React，请考虑使用直接在 HTML 和 JavaScript 中实现的旧的 Fabric UI 组件。 某些示例模板位于 [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) 存储库中。 特别在对话框和导航子文件夹中查看。 示例 [Excel-Add-in-SalesLeads](https://github.com/OfficeDev/Excel-Add-in-SalesLeads) 使用消息横幅。

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [OfficeExtension.Error 对象（Excel JavaScript API）](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Office 常用 API 错误代码](../reference/javascript-api-for-office-error-codes.md)

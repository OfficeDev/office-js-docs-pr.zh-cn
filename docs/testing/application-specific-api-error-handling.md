---
title: 使用特定于应用程序的 JavaScript API 处理错误
description: 了解 Excel、Word、PowerPoint 和其他特定于应用程序的 JavaScript API 错误处理逻辑，以考虑运行时错误。
ms.date: 10/27/2022
ms.localizationpriority: medium
ms.openlocfilehash: 21d8d3eef36f919f95459fd8e0b3037c1d5ae1b1
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/28/2022
ms.locfileid: "68767152"
---
# <a name="error-handling-with-the-application-specific-javascript-apis"></a>使用特定于应用程序的 JavaScript API 处理错误

使用 [特定于应用程序的 Office JavaScript API](../develop/application-specific-api-model.md) 生成外接程序时，请务必包含错误处理逻辑，以考虑运行时错误。 由于 API 的异步性质，这样做至关重要。

## <a name="best-practices"></a>最佳做法

在我们的[代码示例](https://github.com/OfficeDev/Office-Add-in-samples)和[Script Lab](../overview/explore-with-script-lab.md)代码片段中，你会注意到，对 、 `PowerPoint.run`或 `Word.run` 的每个调用`Excel.run`都附带一个`catch`语句来捕获任何错误。 建议在使用特定于应用程序的 API 生成外接程序时使用相同的模式。

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

当 Office JavaScript API 请求未成功运行时，API 将返回包含以下属性的错误对象。

- **代码**：`code`错误消息的 属性包含一个字符串，该字符串是 *{application}* 表示 Excel、PowerPoint 或 Word 的`OfficeExtension.ErrorCodes``{application}.ErrorCodes`一部分或其中。 例如，错误代码“InvalidReference”指示引用对于指定操作无效。 错误代码尚未本地化。

- **消息**：错误消息的 `message` 属性包含本地化字符串中的错误摘要。 错误消息不适合最终用户使用;应使用错误代码和适当的业务逻辑来确定加载项向最终用户显示的错误消息。

- **debugInfo**：出现此信息时，错误消息的 `debugInfo` 属性将提供其他信息，帮助理解错误根本原因。

> [!NOTE]
> 如果使用 `console.log()` 将错误消息打印到控制台，则这些消息仅在服务器上可见。 最终用户不会在加载项任务窗格或 Office 应用程序中的任意位置看到这些错误消息。 若要向用户报告错误，请参阅 [错误通知](#error-notifications)。

## <a name="error-codes-and-messages"></a>错误代码和消息

下表列出了特定于应用程序的 API 可能返回的错误。

> [!NOTE]
> 下表列出了在使用特定于应用程序的 API 时可能会遇到的错误消息。 如果使用的是通用 API，请参阅 [Office 通用 API 错误代码](../reference/javascript-api-for-office-error-codes.md) ，了解相关的错误消息。

|错误代码 | 错误消息 | 备注 |
|:----------|:--------------|:------|
|`AccessDenied` |无法执行所请求的操作。|*没有。* |
|`ActivityLimitReached`|已达到活动限制。|*没有。* |
|`ApiNotAvailable`|请求的 API 不可用。|*没有。* |
|`ApiNotFound`|找不到你尝试使用的 API。 它可能在较新版本的 Office 应用程序中可用。 有关详细信息 [，请参阅 Office 外接程序的 Office 客户端应用程序和平台可用性](/javascript/api/requirement-sets) 。|*没有。* |
|`BadPassword`|提供的密码不正确。|*没有。* |
|`Conflict`|由于冲突，无法处理请求。|*没有。* |
|`ContentLengthRequired`|`Content-length`缺少 HTTP 标头。|*没有。* |
|`GeneralException`|处理请求时出现内部错误。|*没有。* |
|`InsertDeleteConflict`|尝试的插入或删除操作导致冲突。|*没有。* |
|`InvalidArgument` |自变量无效、缺少或格式不正确。|*没有。* |
|`InvalidBinding` |由于之前的更新，此对象绑定不再有效。|*没有。* |
|`InvalidOperation`|尝试的操作对于对象无效。|*没有。* |
|`InvalidReference`|此引用对于当前操作无效。|*没有。* |
|`InvalidRequest`  |无法处理此请求。|*没有。* |
|`InvalidSelection`|当前选定内容对于此操作无效。|*没有。* |
|`ItemAlreadyExists`|所创建的资源已存在。|*没有。* |
|`ItemNotFound` |所请求的资源不存在。|*没有。* |
|`MemoryLimitReached`|已达到内存限制。 无法完成操作。|*没有。* |
|`NotImplemented`|所请求的功能未实现。| 这可能意味着 API 处于预览状态，或者仅在特定平台 (（例如仅联机) ）上受支持。 有关详细信息 [，请参阅 Office 外接程序的 Office 客户端应用程序和平台可用性](/javascript/api/requirement-sets) 。|
|`RequestAborted`|请求在运行时已中止。|*没有。* |
|`RequestPayloadSizeLimitExceeded`|请求有效负载大小已超出限制。 有关详细信息，请参阅 [Office 外接程序的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) 一文。| 此错误仅在 Office web 版 发生。|
|`ResponsePayloadSizeLimitExceeded`|响应有效负载大小已超出限制。 有关详细信息，请参阅 [Office 外接程序的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) 一文。|  此错误仅在 Office web 版 发生。|
|`ServiceNotAvailable`|服务不可用。|*没有。* |
|`Unauthenticated` |所需的身份验证信息缺少或无效。|*没有。* |
|`UnsupportedFeature`|操作失败，因为源工作表包含一个或多个不支持的功能。|*没有。* |
|`UnsupportedOperation`|不支持正在尝试的操作。|*没有。* |

### <a name="excel-specific-error-codes-and-messages"></a>Excel 特定的错误代码和消息

|错误代码 | 错误消息 | 备注 |
|:----------|:--------------|:------|
|`EmptyChartSeries`|尝试的操作失败，因为图表系列为空。|*没有。* |
|`FilteredRangeConflict`|尝试的操作会导致与筛选范围冲突。|*没有。* |
|`FormulaLengthExceedsLimit`|所应用公式的字节码超出了最大长度限制。 对于 32 位计算机上的 Office，字节码长度限制为 16384 个字符。 在 64 位计算机上，字节码长度限制为 32768 个字符。| 此错误在Excel web 版和桌面上都发生。|
|`GeneralException`|*各种。*|数据类型 API 返回 `GeneralException` 带有动态错误消息的错误。 这些消息引用作为错误源的单元格，以及导致错误的问题，例如：“单元格 A1 缺少所需的属性 `type`。”|
|`InactiveWorkbook`|操作失败，因为多个工作簿处于打开状态，并且此 API 调用的工作簿已失去焦点。|*没有。* |
|`InvalidOperationInCellEditMode`|当 Excel 处于“编辑单元格”模式时，该操作不可用。 使用 **Enter** 或 **Tab** 键退出编辑模式，或者选择另一个单元格，然后重试。|*没有。* |
|`MergedRangeConflict`|无法完成操作。 表不能与其他表、数据透视表、查询结果、合并单元格或 XML 映射重叠。|*没有。* |
|`NonBlankCellOffSheet`|Microsoft Excel 无法插入新单元格，因为它会将非空单元格从工作表的末尾推送。 这些非空单元格可能显示为空，但具有空白值、某些格式或公式。 删除足够的行或列，以便为要插入的内容腾出空间，然后重试。|*没有。* |
|`OperationCellsExceedLimit`|尝试的操作影响超过 33554000 个单元格的限制。| `TableColumnCollection.add API`如果 触发此错误，请确认工作表中没有意外数据，但表外没有意外数据。 特别是，检查工作表最右侧列中的数据。 删除意外数据以解决此错误。 验证操作处理多少个单元格的一种方法是运行以下计算： `(number of table rows) x (16383 - (number of table columns))`。 数字 16383 是 Excel 支持的最大列数。 <br><br>此错误仅在 Excel web 版 发生。 |
|`PivotTableRangeConflict`|尝试的操作会导致与数据透视表范围冲突。|*没有。* |
|`RangeExceedsLimit`|区域中的单元格计数已超过支持的最大数目。 有关详细信息，请参阅 [Office 外接程序的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) 一文。|*没有。* |
|`RefreshWorkbookLinksBlocked`|操作失败，因为用户未授予刷新外部工作簿链接的权限。|*没有。* |
|`UnsupportedSheet`|此工作表类型不支持此操作，因为它是宏或图表工作表。|*没有。* |

### <a name="word-specific-error-codes-and-messages"></a>特定于单词的错误代码和消息

|错误代码 | 错误消息 | 备注 |
|:----------|:--------------|:------|
|`SearchDialogIsOpen`|搜索对话框已打开。|*没有。* |
|`SearchStringInvalidOrTooLong`|搜索字符串无效或太长。| 搜索字符串最大为 255 个字符。 |

## <a name="error-notifications"></a>错误通知

向用户报告错误的方式取决于所使用的 UI 系统。 如果使用 React 作为 UI 系统，请使用 Fluent UI 组件和设计元素。 从此 [Fluent UI 页](https://developer.microsoft.com/fluentui#/controls/web)中选择适当的控件。 建议使用消息栏、对话框或模式传达错误消息。 如果错误位于用户的输入中，请在输入控件附近以红色粗体显示错误。 示例 [Office-Add-in-Microsoft-Graph-React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/0706cc67645675a48747f1fec1b1e5722b575b11/Samples/auth/Office-Add-in-Microsoft-Graph-React)使用 MessageBar 元素，并对其进行修改以考虑外接程序任务窗格中的个性菜单。

如果不对 UI 使用 React，请考虑使用在 HTML 和 JavaScript 中直接实现的旧版 Fabric UI 组件。 一些示例模板位于 [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) 存储库中。 尤其在对话框和导航子文件夹中查看。 示例 [Excel-Add-in-SalesLeads](https://github.com/OfficeDev/Excel-Add-in-SalesLeads) 使用消息横幅。

## <a name="see-also"></a>另请参阅

- [OfficeExtension.Error 对象](/javascript/api/office/officeextension.error)
- [Office 常用 API 错误代码](../reference/javascript-api-for-office-error-codes.md)

---
title: JavaScript API Excel错误处理
description: 了解如何Excel JavaScript API 错误处理逻辑，以考虑运行时错误。
ms.date: 01/15/2021
localization_priority: Normal
ms.openlocfilehash: 42ef52b5d20a2c2d1284f57c7b4026ff2c71ebdd
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349509"
---
# <a name="error-handling-with-the-excel-javascript-api"></a>JavaScript API Excel错误处理

使用 Excel JavaScript API 生成加载项时，请务必加入错误处理逻辑，以便解决运行时错误。 鉴于 API 的异步特性，这样做非常关键。

> [!NOTE]
> 有关 JavaScript API 的方法和异步特性 `sync()` Excel，请参阅 Excel 外接程序中的[Office JavaScript 对象模型](excel-add-ins-core-concepts.md)。

## <a name="best-practices"></a>最佳做法

通过本文档中的代码示例，你会注意到每次调用 `Excel.run` 时，都会带上 `catch` 语句，以便捕获 `Excel.run` 内出现的任何错误。 建议在使用 Excel JavaScript API 生成加载项时使用相同模式。

```js
Excel.run(function (context) {
  
  // Excel JavaScript API calls here

  // Await the completion of context.sync() before continuing.
  return context.sync()
    .then(function () {
      console.log("Finished!");
    })
}).catch(errorHandlerFunction);
```

## <a name="api-errors"></a>API 错误

当 Excel JavaScript API 请求无法成功运行时，API 将返回一个包含以下属性的错误对象。

- **代码**：错误消息的 `code` 属性包含一个字符串，它属于 `OfficeExtension.ErrorCodes` 或 `Excel.ErrorCodes` 列表的一部分。 例如，错误代码“InvalidReference”指示引用对于指定操作无效。 错误代码尚未本地化。

- **消息**：错误消息的 `message` 属性包含本地化字符串中的错误摘要。 错误消息并非针对最终用户的使用情况；应使用错误代码和相应的业务逻辑来确定外接程序显示给最终用户的错误消息。

- **debugInfo**：出现此信息时，错误消息的 `debugInfo` 属性将提供其他信息，帮助理解错误根本原因。

> [!NOTE]
> 如果使用 `console.log()` 将错误消息打印到控制台，那么这些消息只会在服务器上可见。 最终用户不会在加载项任务窗格或应用程序中的任何位置看到这些Office消息。

## <a name="error-messages"></a>错误消息

下表是 API 可能返回的错误列表。

|错误代码 | 错误消息 |
|:----------|:--------------|
|`AccessDenied` |无法执行所请求的操作。|
|`ActivityLimitReached`|已达到活动限制。|
|`ApiNotAvailable`|请求的 API 不可用。|
|`ApiNotFound`|找不到您尝试使用的 API。 它在更高版本的 Excel 中Excel。 有关详细信息[，Excel JavaScript API](../reference/requirement-sets/excel-api-requirement-sets.md)要求集文章。|
|`BadPassword`|你提供的密码不正确。|
|`Conflict`|由于冲突，无法处理请求。|
|`ContentLengthRequired`|`Content-length`HTTP 标头缺失。|
|`GeneralException`|处理请求时出现内部错误。|
|`InactiveWorkbook`|操作失败，因为多个工作簿已打开，并且此 API 调用的工作簿已失去焦点。|
|`InsertDeleteConflict`|尝试的插入或删除操作导致冲突。|
|`InvalidArgument` |自变量无效、缺少或格式不正确。|
|`InvalidBinding`  |由于之前的更新，此对象绑定不再有效。|
|`InvalidOperation`|尝试的操作对于对象无效。|
|`InvalidReference`|此引用对于当前操作无效。|
|`InvalidRequest`  |无法处理此请求。|
|`InvalidSelection`|当前选定内容对于此操作无效。|
|`ItemAlreadyExists`|所创建的资源已存在。|
|`ItemNotFound` |所请求的资源不存在。|
|`NonBlankCellOffSheet`|Microsoft Excel无法插入新单元格，因为它会将非空单元格推送到工作表末尾。 这些非空单元格可能为空，但具有空值、某些格式或公式。 删除足够的行或列，为要插入的行或列提供空间，然后重试。|
|`NotImplemented`|所请求的功能未实现。|
|`RangeExceedsLimit`|该范围中的单元格计数已超出支持的最大数。 有关详细信息[，请参阅Office加载项](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)的资源限制和性能优化一文。|
|`RequestAborted`|请求在运行时已中止。|
|`RequestPayloadSizeLimitExceeded`|请求有效负载大小已超出限制。 有关详细信息[，请参阅Office加载项](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)的资源限制和性能优化一文。 <br><br>此错误仅出现在Excel web 版。|
|`ResponsePayloadSizeLimitExceeded`|响应有效负载大小已超出限制。 有关详细信息[，请参阅Office加载项](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)的资源限制和性能优化一文。  <br><br>此错误仅出现在Excel web 版。|
|`ServiceNotAvailable`|服务不可用。|
|`Unauthenticated` |所需的身份验证信息缺少或无效。|
|`UnsupportedOperation`|不支持正在尝试的操作。|
|`UnsupportedSheet`|此工作表类型不支持此操作，因为它是一个宏或图表工作表。|

> [!NOTE]
> 上表列出了使用 JavaScript API 时Excel错误消息。 如果你使用通用 API 而不是特定于应用程序的 JavaScript API Excel，请参阅Office[通用 API](../reference/javascript-api-for-office-error-codes.md)错误代码以了解相关的错误消息。

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [OfficeExtension.Error 对象（Excel JavaScript API）](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Office 常用 API 错误代码](../reference/javascript-api-for-office-error-codes.md)

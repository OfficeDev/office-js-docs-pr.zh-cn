---
title: 错误处理
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: caba29f7d6949cc6d9df1498ac0a3d4f5de6c4ee
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579812"
---
# <a name="error-handling"></a>错误处理

使用 Excel JavaScript API 生成加载项时，请务必加入错误处理逻辑，以便解决运行时错误。鉴于 API 的异步特性，这样做非常关键。

> [!NOTE]
> 若要详细了解 **sync()** 和 Excel JavaScript API 异步特性，请参阅 [Excel JavaScript API 基本编程概念](excel-add-ins-core-concepts.md)。

## <a name="best-practices"></a>最佳做法

通过本文档中的代码示例，你会注意到每次调用 `Excel.run` 时，都会带上 `catch` 语句，以便捕获 `Excel.run` 内出现的任何错误。建议在使用 Excel JavaScript API 生成加载项时使用相同模式。

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

当 Excel JavaScript API 请求无法成功运行时，API 将返回错误对象，其中包含以下属性： 

- **代码**：错误消息的 `code` 属性包含一个字符串，它属于 `OfficeExtension.ErrorCodes` 或 `Excel.ErrorCodes` 列表的一部分。例如，错误代码“InvalidReference”指示引用对于指定操作无效。错误代码尚未本地化。 

- **消息**：错误消息的 `message` 属性包含本地化字符串中的错误摘要。错误消息并非针对最终用户的使用情况；应使用错误代码和相应的业务逻辑来确定加载项显示给最终用户的错误消息。

- **debugInfo**：出现此信息时，错误消息的 `debugInfo` 属性将提供其他信息，帮助理解错误的根本原因。 

> [!NOTE]
> 如果使用 `console.log()` 将错误消息打印到控制台，那么这些消息只会在服务器上可见。最终用户不会在加载项任务窗格或主机应用的其他任何位置看到这些错误消息。

## <a name="error-messages"></a>错误消息

下表定义了 API 可能返回的错误列表。

|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |自变量无效、缺少或格式不正确。|
|InvalidRequest  |无法处理此请求。|
|InvalidReference|此引用对于当前操作无效。|
|InvalidBinding  |由于之前的更新，此对象绑定不再有效。|
|InvalidSelection|当前选定内容对于此操作无效。|
|Unauthenticated |所需的身份验证信息缺少或无效。|
|AccessDenied |无法执行所请求的操作。|
|ItemNotFound |所请求的资源不存在。|
|ActivityLimitReached|已达到活动限制。|
|GeneralException|处理请求时出现内部错误。|
|NotImplemented  |所请求的功能未实现。|
|ServiceNotAvailable|服务不可用。|
|冲突|由于冲突，无法处理请求。|
|ItemAlreadyExists|所创建的资源已存在。|
|UnsupportedOperation|不支持正在尝试的操作。|
|RequestAborted|请求在运行时已中止。|
|wApiNotAvailable|请求的 API 不可用。|
|InsertDeleteConflict|尝试的插入或删除操作导致冲突。|
|InvalidOperation|尝试的操作对于对象无效。|

## <a name="see-also"></a>另请参阅

- [使用 Excel JavaScript API 的基本编程概念](excel-add-ins-core-concepts.md)
- [OfficeExtension.Error 对象 (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)

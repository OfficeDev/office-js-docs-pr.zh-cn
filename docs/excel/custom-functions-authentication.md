---
ms.date: 07/09/2019
description: 使用 Excel 中的自定义函数对用户进行身份验证。
title: 自定义函数的身份验证
localization_priority: Normal
ms.openlocfilehash: aa966aeb8d8161339bab0161b4cc329a9b495d08
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950682"
---
# <a name="authentication-for-custom-functions"></a>自定义函数的身份验证

在某些情况下，你的自定义函数需要对用户进行身份验证才能访问受保护的资源。 虽然自定义函数不需要特定的身份验证方法，但你应该知道自定义函数在与加载项的任务窗格和其他 UI 元素不同的运行时中运行。 因此，你需要使用 `OfficeRuntime.storage` 对象和对话框 API 在两个运行时之间来回传递数据。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="officeruntimestorage-object"></a>OfficeRuntime.storage 对象

自定义函数运行时在全局窗口中没有可用的 `localStorage` 对象，你通常可以在其中存储数据。 相反，你应该使用 [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) 来设置和获取数据，从而在自定义函数和任务窗格之间共享数据。

此外，使用 `storage` 对象也有好处；它使用安全的沙盒环境，以便其他加载项无法访问你的数据。

### <a name="suggested-usage"></a>建议的用法

如果你需要通过任务窗格或自定义函数进行身份验证，请选中 `storage` 以查看是否已获取访问令牌。 如果没有，请使用对话框 API 对用户进行身份验证，检索访问令牌，然后将令牌存储在 `storage` 中以备将来使用。

## <a name="dialog-api"></a>对话框 API

如果令牌不存在，则应使用对话框 API 让用户登录。 用户输入凭据后，生成的访问令牌可以存储在 `storage` 中。

> [!NOTE]
> 自定义函数运行时使用的 Dialog 对象与任务窗格使用的浏览器引擎运行时中的 Dialog 对象略有不同。 它们都称为“对话框 API”，但在自定义函数运行时中使用 `OfficeRuntime.Dialog` 对用户进行身份验证。

有关如何使用 `Dialog` 对象的信息，请参阅[自定义函数对话框](/office/dev/add-ins/excel/custom-functions-dialog)。

在构想整个身份验证过程时，将加载项的任务窗格和 UI 元素以及加载项的自定义函数部分视为可以通过 `OfficeRuntime.storage` 进行相互通信的单独实体，这样做可能对你有所帮助。

下图概述了此基本过程。 请注意，虚线指示虽然它们执行单独的操作，但自定义函数和加载项的任务窗格都是整个加载项的一部分。

1. 你可以从 Excel 工作簿中的单元格发出自定义函数调用。
2. 自定义函数使用 `Dialog` 将你的用户凭据传递给网站。
3. 该网站随后会将访问令牌返回给自定义函数。
4. 然后，自定义函数会将此访问令牌存储在 `storage` 中。
5. 加载项的任务窗格将从 `storage` 访问该令牌。

![自定义函数的关系图，使用对话框 API 获取访问令牌，然后通过 OfficeRuntime API 与任务窗格共享令牌。](../images/authentication-diagram.png "身份验证图。")

## <a name="storing-the-token"></a>存储令牌

以下示例来自[在自定义函数中使用 OfficeRuntime.storage](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) 代码示例。 有关在自定义函数与任务窗格之间共享数据的完整示例，请参阅此代码示例。

如果使用自定义函数进行身份验证，则它会收到访问令牌，并且需要将其存储在 `storage` 中。 以下代码示例演示如何调用 `storage.setItem` 方法来存储值。 `storeValue` 函数是一个自定义函数，例如用于存储用户的值。 你可以对其进行修改以存储所需的任何令牌值。

```js
/**
 * Stores a key-value pair into OfficeRuntime.storage.
 * @customfunction
 * @param {string} key Key of item to put into storage.
 * @param {*} value Value of item to put into storage.
 */
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

当任务窗格需要访问令牌时，它可以从 `storage` 检索令牌。 以下代码示例演示如何使用 `storage.getItem` 方法来检索令牌。

```js
/**
 * Read a token from storage.
 * @customfunction GETTOKEN
 */
function receiveTokenFromCustomFunction() {
  var key = "token";
  var tokenSendStatus = document.getElementById('tokenSendStatus');
  OfficeRuntime.storage.getItem(key).then(function (result) {
     tokenSendStatus.value = "Success: Item with key '" + key + "' read from storage.";
     document.getElementById('tokenTextBox2').value = result;
  }, function (error) {
     tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from storage. " + error;
  });
}
```

## <a name="general-guidance"></a>一般指导

Office 加载项基于 Web，你可以使用任何 Web 身份验证技术。 使用自定义函数实施自己的身份验证时，不必遵循特定的模式或方法。 你可能希望查阅有关各种身份验证模式的文档，请先参阅[这篇关于通过外部服务进行授权的文章](/office/dev/add-ins/develop/auth-external-add-ins)。  

在开发自定义函数时，避免使用以下位置存储数据：  

- `localStorage`：自定义函数无权访问全局 `window` 对象，因此无法访问 `localStorage` 中存储的数据。
- `Office.context.document.settings`：此位置不安全，使用加载项的任何人员都可以提取相关信息。

## <a name="next-steps"></a>后续步骤
了解[自定义函数的对话框 API](custom-functions-dialog.md)。

## <a name="see-also"></a>另请参阅

* [自定义函数体系结构](custom-functions-architecture.md)
* [使用自定义函数接收和处理数据](custom-functions-web-reqs.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [Excel 自定义函数教程](excel-tutorial-custom-functions.md)

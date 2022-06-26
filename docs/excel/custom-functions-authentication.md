---
ms.date: 06/15/2022
description: 使用不使用共享运行时的自定义函数对用户进行身份验证。
title: 对没有共享运行时的自定义函数进行身份验证
ms.localizationpriority: medium
ms.openlocfilehash: 0f4493f9cf68236a9d9d83ebd3299c9ce3371560
ms.sourcegitcommit: d8fbe472b35c758753e5d2e4b905a5973e4f7b52
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/25/2022
ms.locfileid: "66229678"
---
# <a name="authentication-for-custom-functions-without-a-shared-runtime"></a>对没有共享运行时的自定义函数进行身份验证

在某些情况下，不使用共享运行时的自定义函数需要对用户进行身份验证才能访问受保护的资源。 在仅限 JavaScript 的运行时中不使用共享运行时的自定义函数。 因此，如果外接程序具有任务窗格，则需要在仅限 JavaScript 的运行时和任务窗格使用的支持 HTML 的运行时之间来回传递数据。 使用 [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) 对象和特殊对话框 API 执行此操作。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a>OfficeRuntime.storage 对象

仅限 JavaScript 的运行时在全局窗口中没有 `localStorage` 可用的对象，通常会在其中存储数据。 相反，代码应使用 `OfficeRuntime.storage` 设置和获取数据在自定义函数和任务窗格之间共享数据。

### <a name="suggested-usage"></a>建议的用法

当需要从不使用共享运行时的自定义函数外接程序进行身份验证时，代码应检查 `OfficeRuntime.storage` 是否已获取访问令牌。 如果没有，请使用 [OfficeRuntime.displayWebDialog](/javascript/api/office-runtime#office-runtime-officeruntime-displaywebdialog-function(1)) 对用户进行身份验证，检索访问令牌，然后将令牌存储在 `OfficeRuntime.storage` 其中供将来使用。

## <a name="dialog-api"></a>对话框 API

如果某个令牌不存在，则应使用 `OfficeRuntime.dialog` API 要求用户登录。 用户输入凭据后，生成的访问令牌可以存储为项 `OfficeRuntime.storage`。

> [!NOTE]
> 仅限 JavaScript 的运行时使用与任务窗格使用的浏览器引擎运行时中的对话框对象略有不同的对话对象。 它们都称为“对话框 API”，但使用 [OfficeRuntime.displayWebDialog](/javascript/api/office-runtime#office-runtime-officeruntime-displaywebdialog-function(1)) 在仅限 JavaScript 的运行时（*而不是* [Office.ui.displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1))）中对用户进行身份验证。

下图概述了此基本过程。 虚线表示自定义函数和外接程序的任务窗格都是加载项整体的一部分，尽管它们使用单独的运行时。

1. 你可以从 Excel 工作簿中的单元格发出自定义函数调用。
2. 自定义函数使用 `OfficeRuntime.dialog` 将你的用户凭据传递给网站。
3. 该网站随后会将访问令牌返回给自定义函数。
4. 然后，自定义函数将此访问令牌设置为其中 `OfficeRuntime.storage`的项。
5. 加载项的任务窗格将从 `OfficeRuntime.storage` 访问该令牌。

![使用对话框 API 获取访问令牌，然后通过 OfficeRuntime.storage API 与任务窗格共享令牌的自定义函数示意图。](../images/authentication-diagram.png "身份验证图。")

## <a name="storing-the-token"></a>存储令牌

以下示例来自[在自定义函数中使用 OfficeRuntime.storage](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/AsyncStorage) 代码示例。 有关在不使用共享运行时的加载项中的自定义函数和任务窗格之间共享数据的完整示例，请参阅此代码示例。

如果使用自定义函数进行身份验证，则它会收到访问令牌，并且需要将其存储在 `OfficeRuntime.storage` 中。 以下代码示例演示如何调用 `storage.setItem` 方法来存储值。 该 `storeValue` 函数是一个自定义函数，用于存储用户的值。 你可以对其进行修改以存储所需的任何令牌值。

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

当任务窗格需要访问令牌时，它可以从 `OfficeRuntime.storage` 项中检索令牌。 以下代码示例演示如何使用 `storage.getItem` 方法来检索令牌。

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

Office 加载项基于 Web，你可以使用任何 Web 身份验证技术。 使用自定义函数实施自己的身份验证时，不必遵循特定的模式或方法。 你可能希望查阅有关各种身份验证模式的文档，请先参阅[这篇关于通过外部服务进行授权的文章](../develop/auth-external-add-ins.md)。  

在开发自定义函数时，避免使用以下位置存储数据：

- `localStorage`：不使用共享运行时的自定义函数无权访问全局 `window` 对象，因此无法访问存储在其中 `localStorage`的数据。
- `Office.context.document.settings`：此位置不安全，任何使用加载项的人都可以提取信息。

## <a name="dialog-box-api-example"></a>对话框 API 示例

在以下代码示例中，函 `getTokenViaDialog` 数使用该 `OfficeRuntime.displayWebDialog` 函数显示对话框。 提供此示例是为了显示方法的功能，而不是演示如何进行身份验证。

```JavaScript
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this isn't a sufficient example of authentication but is intended to show the capabilities of the displayWebDialog method.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once. Wait for previous dialog box's token.
      let timeout = 5;
      let count = 0;
      var intervalId = setInterval(function () {
        count++;
        if(_cachedToken) {
          resolve(_cachedToken);
          clearInterval(intervalId);
        }
        if(count >= timeout) {
          reject("Timeout while waiting for token");
          clearInterval(intervalId);
        }
      }, 1000);
    } else {
      _dialogOpen = true;
      OfficeRuntime.displayWebDialog(url, {
        height: '50%',
        width: '50%',
        onMessage: function (message, dialog) {
          _cachedToken = message;
          resolve(message);
          dialog.close();
          return;
        },
        onRuntimeError: function(error, dialog) {
          reject(error);
        },
      }).catch(function (e) {
        reject(e);
      });
    }
  });
}
```

## <a name="next-steps"></a>后续步骤

了解如何 [调试自定义函数](custom-functions-debugging.md)。

## <a name="see-also"></a>另请参阅

* [自定义函数的仅限 JavaScript 的运行时](custom-functions-runtime.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
---
ms.date: 05/17/2020
description: 在不使用任务窗格Excel自定义函数对用户进行身份验证。
title: 无 UI 自定义函数的身份验证
localization_priority: Normal
ms.openlocfilehash: 94eadd343f969e6dbd83881764fac936acf0704b
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937142"
---
# <a name="authentication-for-ui-less-custom-functions"></a>无 UI 自定义函数的身份验证

在某些情况下，不使用任务窗格或其他用户界面元素的自定义函数 (无 UI 自定义函数) 将需要对用户进行身份验证，才能访问受保护的资源。 请注意，无 UI 自定义函数在仅 JavaScript 运行时中运行。 因此，你需要在仅 JavaScript 运行时和使用 对象和对话框 API 的外接程序使用的典型浏览器引擎运行时之间来回 `OfficeRuntime.storage` 传递数据。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a>OfficeRuntime.storage 对象

无 UI 自定义函数使用的仅 JavaScript 运行时在全局窗口（通常存储数据）上 `localStorage` 没有可用的对象。 相反，您应使用 [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) 设置和获取数据，在无 UI 自定义函数和任务窗格之间共享数据。

### <a name="suggested-usage"></a>建议的用法

当你需要从无 UI 自定义函数进行身份验证时，请检查 `storage` 是否获取了访问令牌。 如果没有，请使用对话框 API 对用户进行身份验证，检索访问令牌，然后将令牌存储在 `storage` 中以备将来使用。

## <a name="dialog-api"></a>对话框 API

如果令牌不存在，则应使用对话框 API 让用户登录。 用户输入凭据后，生成的访问令牌可以存储在 `storage` 中。

> [!NOTE]
> 仅 JavaScript 运行时使用的 Dialog 对象与任务窗格使用的浏览器引擎运行时中的 Dialog 对象略有不同。 它们都称为"对话框 API"，但用于对仅 `OfficeRuntime.Dialog` JavaScript 运行时中的用户进行身份验证。

下图概述了此基本过程。 虚线表示无 UI 自定义函数和加载项的任务窗格都是整个加载项的一部分，尽管它们使用单独的运行时。

1. 从工作簿中的单元格发出无 UI 自定义函数Excel调用。
2. 无 UI 自定义函数用于 `Dialog` 将用户凭据传递到网站。
3. 然后，此网站会向无 UI 自定义函数返回访问令牌。
4. 然后，无 UI 自定义函数将此访问令牌设置到 `storage` 。
5. 加载项的任务窗格将从 `storage` 访问该令牌。

![使用对话框 API 获取访问令牌，然后通过 OfficeRuntime.storage API 与任务窗格共享令牌的自定义函数关系图。](../images/authentication-diagram.png "身份验证图表。")

## <a name="storing-the-token"></a>存储令牌

以下示例来自[在自定义函数中使用 OfficeRuntime.storage](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) 代码示例。 有关在无 UI 自定义函数和任务窗格之间共享数据的完整示例，请参阅此代码示例。

如果无 UI 自定义函数进行身份验证，则它会收到访问令牌，并且将需要将令牌存储在 中 `storage` 。 以下代码示例演示如何调用 `storage.setItem` 方法来存储值。 `storeValue`函数是无 UI 的自定义函数，例如用于存储用户的值。 你可以对其进行修改以存储所需的任何令牌值。

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

Office 加载项基于 Web，你可以使用任何 Web 身份验证技术。 没有特定的模式或方法必须遵循，以使用无 UI 自定义函数实现你自己的身份验证。 你可能希望查阅有关各种身份验证模式的文档，请先参阅[这篇关于通过外部服务进行授权的文章](../develop/auth-external-add-ins.md)。  

在开发自定义函数时，请避免使用下列位置来存储数据：。

- `localStorage`：无 UI 的自定义函数无法访问全局对象，因此无法访问 `window` 中存储的数据 `localStorage` 。
- `Office.context.document.settings`：此位置不安全，使用加载项的任何人员都可以提取相关信息。

## <a name="dialog-box-api-example"></a>对话框 API 示例

在下面的代码示例中，函数使用 `getTokenViaDialog` API `Dialog` `displayWebDialogOptions` 函数显示对话框。 提供此示例是显示对象的功能 `Dialog` ，而不是演示如何进行身份验证。

```JavaScript
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
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
了解如何调试 [无 UI 自定义函数](custom-functions-debugging.md)。

## <a name="see-also"></a>另请参阅

* [无 UI 的运行时Excel自定义函数](custom-functions-runtime.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
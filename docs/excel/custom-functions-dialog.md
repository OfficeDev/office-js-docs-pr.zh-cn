---
ms.date: 05/06/2019
description: 在 Excel 中使用 JavaScript 通过自定义函数创建对话框。
title: 通过自定义函数显示对话框
localization_priority: Priority
ms.openlocfilehash: 3d7a657402c319b2394c7331b69314b2e5591890
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628141"
---
# <a name="display-a-dialog-box-from-a-custom-function"></a>通过自定义函数显示对话框

如果你的自定义函数需要与用户进行交互，可以使用[`Office.Dialog`对象](/javascript/api/office-runtime/officeruntime.dialog?view=office-js)创建对话框。 使用该对话框的常见方案是对用户进行身份验证，以便你的自定义函数可以访问 Web 服务。 有关使用自定义函数进行身份验证的详细信息，请参阅[自定义函数身份验证](./custom-functions-authentication.md)。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!NOTE]
> `Office.Dialog` 对象是自定义函数运行时的一部分。 任务窗格不使用 `Dialog` 对象。 若要从任务窗格创建对话框，请参阅[对话框 API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins)。

## <a name="dialog-box-api-example"></a>对话框 API 示例

在下面的代码示例中，函数 `getTokenViaDialog` 使用 `Dialog`API 的 `displayWebDialogOptions` 函数来显示对话框。

```js
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once, wait for previous dialog box's token
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
      Office.displayWebDialogOptions(url, {
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
了解如何[让自定义函数与 XLL 用户定义的函数兼容](make-custom-functions-compatible-with-xll-udf.md)。

## <a name="see-also"></a>另请参阅

* [自定义函数身份验证](custom-functions-authentication.md)
* [使用自定义函数接收和处理数据](custom-functions-web-reqs.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)

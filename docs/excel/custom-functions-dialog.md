---
ms.date: 03/21/2019
description: 在 Excel 中使用 JavaScript 通过自定义函数创建对话框。
title: 自定义函数对话框（预览版）
localization_priority: Priority
ms.openlocfilehash: 0f596825a7a32525a68ef45656f1390196146706
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449258"
---
# <a name="display-a-dialog-box-in-custom-functions"></a>在自定义函数中显示对话框

如果你的自定义函数需要与用户进行交互，可以使用 `OfficeRuntime.Dialog` 对象创建对话框。 使用该对话框的常见方案是对用户进行身份验证，以便你的自定义函数可以访问 Web 服务。 有关使用自定义函数进行身份验证的详细信息，请参阅[自定义函数身份验证](./custom-functions-authentication.md)。

注意：`OfficeRuntime.Dialog` 对象是自定义函数运行时的一部分。 它不能在任务窗格的上下文中使用。 若要从任务窗格创建对话框，请参阅[对话框 API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins)。

## <a name="dialog-api-example"></a>对话框 API 示例

在下面的代码示例中，函数 `getTokenViaDialog` 使用对话框 API 的 `displayWebDialog` 函数来显示对话框。

```js
// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"

function getStock (ticker) {
  return new Promise(function (resolve, reject) {
    // Get a token
    getToken("https://www.contoso.com/auth")
    .then(function (token) {

      // Use token to get stock price
      fetch("https://www.contoso.com/?token=token&ticker= + ticker")
      .then(function (result) {

        // Return stock price to cell
        resolve(result);
      });
    })
    .catch(function (error) {
      reject(error);
    });
  });
  
  //Helper
  function getToken(url) {
    return new Promise(function (resolve,reject) {
      if(_cachedToken) {
        resolve(_cachedToken);
      } else {
        getTokenViaDialog(url)
        .then(function (result) {
          resolve(result);
        })
        .catch(function (result) {
          reject(result);
        });
      }
    });
  }

  function getTokenViaDialog(url) {
    return new Promise (function (resolve, reject) {
      if (_dialogOpen) {
        // Can only have one dialog open at once, wait for previous dialog's token
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
}
```

## <a name="see-also"></a>另请参阅

* [自定义函数元数据](custom-functions-json.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [自定义函数更改日志](custom-functions-changelog.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)

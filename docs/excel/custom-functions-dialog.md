---
ms.date: 03/21/2019
description: 在 Excel 中使用 JavaScript 通过自定义函数创建对话框。
title: 自定义函数对话框（预览版）
localization_priority: Priority
ms.openlocfilehash: 0f596825a7a32525a68ef45656f1390196146706
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30926648"
---
# <a name="display-a-dialog-box-in-custom-functions"></a><span data-ttu-id="2fffb-103">在自定义函数中显示对话框</span><span class="sxs-lookup"><span data-stu-id="2fffb-103">Display a dialog box in custom functions</span></span>

<span data-ttu-id="2fffb-104">如果你的自定义函数需要与用户进行交互，可以使用 `OfficeRuntime.Dialog` 对象创建对话框。</span><span class="sxs-lookup"><span data-stu-id="2fffb-104">If your custom function needs to interact with the user, you can create a dialog box using the `OfficeRuntime.Dialog` object.</span></span> <span data-ttu-id="2fffb-105">使用该对话框的常见方案是对用户进行身份验证，以便你的自定义函数可以访问 Web 服务。</span><span class="sxs-lookup"><span data-stu-id="2fffb-105">A common scenario for using the dialog box is to authenticate a user so that your custom function can access a web service.</span></span> <span data-ttu-id="2fffb-106">有关使用自定义函数进行身份验证的详细信息，请参阅[自定义函数身份验证](./custom-functions-authentication.md)。</span><span class="sxs-lookup"><span data-stu-id="2fffb-106">For more details about authentication with custom functions, see [Custom functions authentication](./custom-functions-authentication.md).</span></span>

<span data-ttu-id="2fffb-107">注意：`OfficeRuntime.Dialog` 对象是自定义函数运行时的一部分。</span><span class="sxs-lookup"><span data-stu-id="2fffb-107">Note: The `OfficeRuntime.Dialog` object is part of the custom functions runtime.</span></span> <span data-ttu-id="2fffb-108">它不能在任务窗格的上下文中使用。</span><span class="sxs-lookup"><span data-stu-id="2fffb-108">It cannot be used from the context of a task pane.</span></span> <span data-ttu-id="2fffb-109">若要从任务窗格创建对话框，请参阅[对话框 API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="2fffb-109">To create a dialog from a task pane, see [Dialog API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).</span></span>

## <a name="dialog-api-example"></a><span data-ttu-id="2fffb-110">对话框 API 示例</span><span class="sxs-lookup"><span data-stu-id="2fffb-110">Dialog API example</span></span>

<span data-ttu-id="2fffb-111">在下面的代码示例中，函数 `getTokenViaDialog` 使用对话框 API 的 `displayWebDialog` 函数来显示对话框。</span><span class="sxs-lookup"><span data-stu-id="2fffb-111">In the following code sample, the function `getTokenViaDialog` uses the Dialog API’s `displayWebDialog` function to display a dialog box.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="2fffb-112">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2fffb-112">See also</span></span>

* [<span data-ttu-id="2fffb-113">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="2fffb-113">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="2fffb-114">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="2fffb-114">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="2fffb-115">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="2fffb-115">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="2fffb-116">自定义函数更改日志</span><span class="sxs-lookup"><span data-stu-id="2fffb-116">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="2fffb-117">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="2fffb-117">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)

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
# <a name="display-a-dialog-box-from-a-custom-function"></a><span data-ttu-id="5b785-103">通过自定义函数显示对话框</span><span class="sxs-lookup"><span data-stu-id="5b785-103">Display a dialog box from a custom function</span></span>

<span data-ttu-id="5b785-104">如果你的自定义函数需要与用户进行交互，可以使用[`Office.Dialog`对象](/javascript/api/office-runtime/officeruntime.dialog?view=office-js)创建对话框。</span><span class="sxs-lookup"><span data-stu-id="5b785-104">If your custom function needs to interact with the user, you can create a dialog box using the `Office.Dialog` object.</span></span> <span data-ttu-id="5b785-105">使用该对话框的常见方案是对用户进行身份验证，以便你的自定义函数可以访问 Web 服务。</span><span class="sxs-lookup"><span data-stu-id="5b785-105">A common scenario for using the dialog box is to authenticate a user so that your custom function can access a web service.</span></span> <span data-ttu-id="5b785-106">有关使用自定义函数进行身份验证的详细信息，请参阅[自定义函数身份验证](./custom-functions-authentication.md)。</span><span class="sxs-lookup"><span data-stu-id="5b785-106">For more details about authentication with custom functions, see [Custom functions authentication](./custom-functions-authentication.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!NOTE]
> <span data-ttu-id="5b785-107">`Office.Dialog` 对象是自定义函数运行时的一部分。</span><span class="sxs-lookup"><span data-stu-id="5b785-107">Note: The `Office.Dialog` object is part of the custom functions runtime.</span></span> <span data-ttu-id="5b785-108">任务窗格不使用 `Dialog` 对象。</span><span class="sxs-lookup"><span data-stu-id="5b785-108">Task panes don't use the `Dialog` object.</span></span> <span data-ttu-id="5b785-109">若要从任务窗格创建对话框，请参阅[对话框 API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="5b785-109">To create a dialog from a task pane, see [Dialog API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).</span></span>

## <a name="dialog-box-api-example"></a><span data-ttu-id="5b785-110">对话框 API 示例</span><span class="sxs-lookup"><span data-stu-id="5b785-110">Dialog API example</span></span>

<span data-ttu-id="5b785-111">在下面的代码示例中，函数 `getTokenViaDialog` 使用 `Dialog`API 的 `displayWebDialogOptions` 函数来显示对话框。</span><span class="sxs-lookup"><span data-stu-id="5b785-111">In the following code sample, the function `getTokenViaDialog` uses the Dialog API’s `Dialog` function to display a dialog box.</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="5b785-112">后续步骤</span><span class="sxs-lookup"><span data-stu-id="5b785-112">Next steps</span></span>
<span data-ttu-id="5b785-113">了解如何[让自定义函数与 XLL 用户定义的函数兼容](make-custom-functions-compatible-with-xll-udf.md)。</span><span class="sxs-lookup"><span data-stu-id="5b785-113">Learn how to [make your custom functions compatible with XLL user-defined functions](make-custom-functions-compatible-with-xll-udf.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="5b785-114">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5b785-114">See also</span></span>

* <span data-ttu-id="5b785-115">[自定义函数身份验证](custom-functions-authentication.md)</span><span class="sxs-lookup"><span data-stu-id="5b785-115">For more information, see [Custom functions authentication](custom-functions-authentication.md).</span></span>
* [<span data-ttu-id="5b785-116">使用自定义函数接收和处理数据</span><span class="sxs-lookup"><span data-stu-id="5b785-116">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="5b785-117">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="5b785-117">Create custom functions in Excel</span></span>](custom-functions-overview.md)

---
ms.date: 05/17/2020
description: 在 Excel 中使用不使用任务窗格的自定义函数对用户进行身份验证。
title: UI 的身份验证-更少的自定义函数
localization_priority: Normal
ms.openlocfilehash: 93073fb23f3f4d30c36faf4927a3aebdafbc887d
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278376"
---
# <a name="authentication-for-ui-less-custom-functions"></a><span data-ttu-id="40151-103">UI 的身份验证-更少的自定义函数</span><span class="sxs-lookup"><span data-stu-id="40151-103">Authentication for UI-less custom functions</span></span>

<span data-ttu-id="40151-104">在某些情况下，不使用任务窗格或其他用户界面元素（无 UI 自定义函数）的自定义函数将需要对用户进行身份验证，以便访问受保护的资源。</span><span class="sxs-lookup"><span data-stu-id="40151-104">In some scenarios your custom function that does not use a task pane or other user interface elements (UI-less custom function) will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="40151-105">请注意，不支持 UI 的自定义函数在仅 JavaScript 运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="40151-105">Be aware that UI-less custom functions run in a JavaScript-only runtime.</span></span> <span data-ttu-id="40151-106">因此，您需要在仅 JavaScript 运行时和使用 `OfficeRuntime.storage` 对象和对话框 API 的大多数加载项使用的典型浏览器引擎运行时之间来回传递数据。</span><span class="sxs-lookup"><span data-stu-id="40151-106">Because of this, you'll need to pass data back and forth between the JavaScript-only runtime and the typical browser engine runtime used by most add-ins using the `OfficeRuntime.storage` object and the Dialog API.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a><span data-ttu-id="40151-107">OfficeRuntime.storage 对象</span><span class="sxs-lookup"><span data-stu-id="40151-107">OfficeRuntime.storage object</span></span>

<span data-ttu-id="40151-108">仅限 UI 的自定义函数使用的仅限 JavaScript 运行时 `localStorage` ，全局窗口中没有可用的对象，通常存储数据。</span><span class="sxs-lookup"><span data-stu-id="40151-108">The JavaScript-only runtime used by UI-less custom functions doesn't have a `localStorage` object available on the global window, where you typically store data.</span></span> <span data-ttu-id="40151-109">相反，应使用[OfficeRuntime](/javascript/api/office-runtime/officeruntime.storage)来设置和获取数据，从而在无 UI 的自定义函数和任务窗格之间共享数据。</span><span class="sxs-lookup"><span data-stu-id="40151-109">Instead, you should share data between UI-less custom functions and task panes by using [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) to set and get data.</span></span>

### <a name="suggested-usage"></a><span data-ttu-id="40151-110">建议的用法</span><span class="sxs-lookup"><span data-stu-id="40151-110">Suggested usage</span></span>

<span data-ttu-id="40151-111">如果需要从无 UI 的自定义函数进行身份验证，请检查是否 `storage` 已获取访问令牌。</span><span class="sxs-lookup"><span data-stu-id="40151-111">When you need to authenticate from a UI-less custom function, check `storage` to see if the access token was already acquired.</span></span> <span data-ttu-id="40151-112">如果没有，请使用对话框 API 对用户进行身份验证，检索访问令牌，然后将令牌存储在 `storage` 中以备将来使用。</span><span class="sxs-lookup"><span data-stu-id="40151-112">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `storage` for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="40151-113">对话框 API</span><span class="sxs-lookup"><span data-stu-id="40151-113">Dialog API</span></span>

<span data-ttu-id="40151-114">如果令牌不存在，则应使用对话框 API 让用户登录。</span><span class="sxs-lookup"><span data-stu-id="40151-114">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="40151-115">用户输入凭据后，生成的访问令牌可以存储在 `storage` 中。</span><span class="sxs-lookup"><span data-stu-id="40151-115">After a user enters their credentials, the resulting access token can be stored in `storage`.</span></span>

> [!NOTE]
> <span data-ttu-id="40151-116">仅 JavaScript 运行时使用与任务窗格使用的浏览器引擎运行时中的 Dialog 对象略有不同的 Dialog 对象。</span><span class="sxs-lookup"><span data-stu-id="40151-116">The JavaScript-only runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes.</span></span> <span data-ttu-id="40151-117">它们都称为 "对话框 API"，但用于 `OfficeRuntime.Dialog` 在仅 JavaScript 运行时中对用户进行身份验证。</span><span class="sxs-lookup"><span data-stu-id="40151-117">They're both referred to as the "Dialog API", but use `OfficeRuntime.Dialog` to authenticate users in the JavaScript-only runtime.</span></span>

<span data-ttu-id="40151-118">下图概述了此基本过程。</span><span class="sxs-lookup"><span data-stu-id="40151-118">The following diagram outlines this basic process.</span></span> <span data-ttu-id="40151-119">点线表示不包含 UI 的自定义函数，外接程序的任务窗格是外接程序的一部分，尽管它们使用单独的运行时。</span><span class="sxs-lookup"><span data-stu-id="40151-119">The dotted line indicates that UI-less custom functions and your add-in's task pane are both part of your add-in as a whole, though they use separate runtimes.</span></span>

1. <span data-ttu-id="40151-120">您从 Excel 工作簿中的单元格发出不带 UI 的自定义函数调用。</span><span class="sxs-lookup"><span data-stu-id="40151-120">You issue a UI-less custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="40151-121">无 UI 的自定义函数用于将 `Dialog` 您的用户凭据传递到网站。</span><span class="sxs-lookup"><span data-stu-id="40151-121">The UI-less custom function uses `Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="40151-122">然后，此网站将访问令牌返回到无 UI 的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="40151-122">This website then returns an access token to the UI-less custom function.</span></span>
4. <span data-ttu-id="40151-123">不带 UI 的自定义函数，然后将此访问令牌设置为 `storage` 。</span><span class="sxs-lookup"><span data-stu-id="40151-123">Your UI-less custom function then sets this access token to the `storage`.</span></span>
5. <span data-ttu-id="40151-124">加载项的任务窗格将从 `storage` 访问该令牌。</span><span class="sxs-lookup"><span data-stu-id="40151-124">Your add-in's task pane accesses the token from `storage`.</span></span>

<span data-ttu-id="40151-125">![自定义函数的关系图，使用对话框 API 获取访问令牌，然后通过 OfficeRuntime API 与任务窗格共享令牌。](../images/authentication-diagram.png "身份验证图。")</span><span class="sxs-lookup"><span data-stu-id="40151-125">![Diagram of custom function using dialog API to get access token, and then share token with task pane through the OfficeRuntime.storage API.](../images/authentication-diagram.png "Authentication diagram.")</span></span>

## <a name="storing-the-token"></a><span data-ttu-id="40151-126">存储令牌</span><span class="sxs-lookup"><span data-stu-id="40151-126">Storing the token</span></span>

<span data-ttu-id="40151-127">以下示例来自[在自定义函数中使用 OfficeRuntime.storage](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) 代码示例。</span><span class="sxs-lookup"><span data-stu-id="40151-127">The following examples are from the [Using OfficeRuntime.storage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample.</span></span> <span data-ttu-id="40151-128">有关在无 UI 的自定义函数和任务窗格之间共享数据的完整示例，请参阅此代码示例。</span><span class="sxs-lookup"><span data-stu-id="40151-128">Refer to this code sample for a complete example of sharing data between UI-less custom functions and the task pane.</span></span>

<span data-ttu-id="40151-129">如果不带 UI 的自定义函数进行身份验证，则它将接收访问令牌，并需要将其存储在中 `storage` 。</span><span class="sxs-lookup"><span data-stu-id="40151-129">If the UI-less custom function authenticates, then it receives the access token and will need to store it in `storage`.</span></span> <span data-ttu-id="40151-130">以下代码示例演示如何调用 `storage.setItem` 方法来存储值。</span><span class="sxs-lookup"><span data-stu-id="40151-130">The following code sample shows how to call the `storage.setItem` method to store a value.</span></span> <span data-ttu-id="40151-131">`storeValue`函数是无 UI 的自定义函数，例如，用于存储来自用户的值。</span><span class="sxs-lookup"><span data-stu-id="40151-131">The `storeValue` function is a UI-less custom function that for example purposes stores a value from the user.</span></span> <span data-ttu-id="40151-132">你可以对其进行修改以存储所需的任何令牌值。</span><span class="sxs-lookup"><span data-stu-id="40151-132">You can modify this to store any token value you need.</span></span>

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

<span data-ttu-id="40151-133">当任务窗格需要访问令牌时，它可以从 `storage` 检索令牌。</span><span class="sxs-lookup"><span data-stu-id="40151-133">When the task pane needs the access token, it can retrieve the token from `storage`.</span></span> <span data-ttu-id="40151-134">以下代码示例演示如何使用 `storage.getItem` 方法来检索令牌。</span><span class="sxs-lookup"><span data-stu-id="40151-134">The following code sample shows how to use the `storage.getItem` method to retrieve the token.</span></span>

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

## <a name="general-guidance"></a><span data-ttu-id="40151-135">一般指导</span><span class="sxs-lookup"><span data-stu-id="40151-135">General guidance</span></span>

<span data-ttu-id="40151-136">Office 加载项基于 Web，你可以使用任何 Web 身份验证技术。</span><span class="sxs-lookup"><span data-stu-id="40151-136">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="40151-137">在使用无 UI 的自定义函数实现自己的身份验证时，必须遵循任何特定的模式或方法。</span><span class="sxs-lookup"><span data-stu-id="40151-137">There is no particular pattern or method you must follow to implement your own authentication with UI-less custom functions.</span></span> <span data-ttu-id="40151-138">你可能希望查阅有关各种身份验证模式的文档，请先参阅[这篇关于通过外部服务进行授权的文章](../develop/auth-external-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="40151-138">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](../develop/auth-external-add-ins.md).</span></span>  

<span data-ttu-id="40151-139">在开发自定义函数时，避免使用以下位置存储数据：</span><span class="sxs-lookup"><span data-stu-id="40151-139">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="40151-140">`localStorage`：不带 UI 的自定义函数不具有对全局对象的访问权限 `window` ，因此无法访问存储在中的数据 `localStorage` 。</span><span class="sxs-lookup"><span data-stu-id="40151-140">`localStorage`: UI-less custom functions do not have access to the global `window` object and therefore have no access to data stored in `localStorage`.</span></span>
- <span data-ttu-id="40151-141">`Office.context.document.settings`：此位置不安全，使用加载项的任何人员都可以提取相关信息。</span><span class="sxs-lookup"><span data-stu-id="40151-141">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the add-in.</span></span>

## <a name="dialog-box-api-example"></a><span data-ttu-id="40151-142">对话框 API 示例</span><span class="sxs-lookup"><span data-stu-id="40151-142">Dialog box API example</span></span>

<span data-ttu-id="40151-143">在下面的代码示例中，函数 `getTokenViaDialog` 使用 `Dialog` API 的 `displayWebDialogOptions` 函数显示一个对话框。</span><span class="sxs-lookup"><span data-stu-id="40151-143">In the following code sample, the function `getTokenViaDialog` uses the `Dialog` API's `displayWebDialogOptions` function to display a dialog box.</span></span> <span data-ttu-id="40151-144">提供此示例以显示对象的功能 `Dialog` ，而不演示如何进行身份验证。</span><span class="sxs-lookup"><span data-stu-id="40151-144">This sample is provided to show the capabilities of the `Dialog` object, not demonstrate how to authenticate.</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="40151-145">后续步骤</span><span class="sxs-lookup"><span data-stu-id="40151-145">Next steps</span></span>
<span data-ttu-id="40151-146">了解如何[调试不带 UI 的自定义函数](custom-functions-debugging.md)。</span><span class="sxs-lookup"><span data-stu-id="40151-146">Learn how to [debug UI-less custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="40151-147">另请参阅</span><span class="sxs-lookup"><span data-stu-id="40151-147">See also</span></span>

* [<span data-ttu-id="40151-148">不带 UI 的 Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="40151-148">Runtime for UI-less Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="40151-149">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="40151-149">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)

---
ms.date: 07/09/2019
description: 使用 Excel 中的自定义函数对用户进行身份验证。
title: 自定义函数的身份验证
localization_priority: Priority
ms.openlocfilehash: f746947122da7ef3d54a0dd3b4f90dd059e5830f
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268136"
---
# <a name="authentication-for-custom-functions"></a><span data-ttu-id="f0a40-103">自定义函数的身份验证</span><span class="sxs-lookup"><span data-stu-id="f0a40-103">Authentication for custom functions</span></span>

<span data-ttu-id="f0a40-104">在某些情况下，你的自定义函数需要对用户进行身份验证才能访问受保护的资源。</span><span class="sxs-lookup"><span data-stu-id="f0a40-104">In some scenarios your custom function will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="f0a40-105">虽然自定义函数不需要特定的身份验证方法，但你应该知道自定义函数在与加载项的任务窗格和其他 UI 元素不同的运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="f0a40-105">While custom functions don't require a specific method of authentication, you should be aware that custom functions run in a separate runtime from the task pane and other UI elements of your add-in.</span></span> <span data-ttu-id="f0a40-106">因此，你需要使用 `OfficeRuntime.storage` 对象和对话框 API 在两个运行时之间来回传递数据。</span><span class="sxs-lookup"><span data-stu-id="f0a40-106">Because of this, you'll need to pass data back and forth between the two runtimes using the `OfficeRuntime.storage` object and the Dialog API.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="officeruntimestorage-object"></a><span data-ttu-id="f0a40-107">OfficeRuntime.storage 对象</span><span class="sxs-lookup"><span data-stu-id="f0a40-107">OfficeRuntime.storage object</span></span>

<span data-ttu-id="f0a40-108">自定义函数运行时在全局窗口中没有可用的 `localStorage` 对象，你通常可以在其中存储数据。</span><span class="sxs-lookup"><span data-stu-id="f0a40-108">The custom functions runtime doesn't have a `localStorage` object available on the global window, where you might typically store data.</span></span> <span data-ttu-id="f0a40-109">相反，你应该使用 [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) 来设置和获取数据，从而在自定义函数和任务窗格之间共享数据。</span><span class="sxs-lookup"><span data-stu-id="f0a40-109">Instead, you should share data between custom functions and task panes by using [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) to set and get data.</span></span>

<span data-ttu-id="f0a40-110">此外，使用 `storage` 对象也有好处；它使用安全的沙盒环境，以便其他加载项无法访问你的数据。</span><span class="sxs-lookup"><span data-stu-id="f0a40-110">Additionally, there is a benefit to using the `storage` object; it uses a secure sandbox environment so that your data cannot be accessed by other add-ins.</span></span>

### <a name="suggested-usage"></a><span data-ttu-id="f0a40-111">建议的用法</span><span class="sxs-lookup"><span data-stu-id="f0a40-111">Suggested usage</span></span>

<span data-ttu-id="f0a40-112">如果你需要通过任务窗格或自定义函数进行身份验证，请选中 `storage` 以查看是否已获取访问令牌。</span><span class="sxs-lookup"><span data-stu-id="f0a40-112">When you need to authenticate either from the task pane or a custom function, check `storage` to see if the access token was already acquired.</span></span> <span data-ttu-id="f0a40-113">如果没有，请使用对话框 API 对用户进行身份验证，检索访问令牌，然后将令牌存储在 `storage` 中以备将来使用。</span><span class="sxs-lookup"><span data-stu-id="f0a40-113">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `storage` for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="f0a40-114">对话框 API</span><span class="sxs-lookup"><span data-stu-id="f0a40-114">Dialog API example</span></span>

<span data-ttu-id="f0a40-115">如果令牌不存在，则应使用对话框 API 让用户登录。</span><span class="sxs-lookup"><span data-stu-id="f0a40-115">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="f0a40-116">用户输入凭据后，生成的访问令牌可以存储在 `storage` 中。</span><span class="sxs-lookup"><span data-stu-id="f0a40-116">After a user enters their credentials, the resulting access token can be stored in `storage`.</span></span>

> [!NOTE]
> <span data-ttu-id="f0a40-117">自定义函数运行时使用的 Dialog 对象与任务窗格使用的浏览器引擎运行时中的 Dialog 对象略有不同。</span><span class="sxs-lookup"><span data-stu-id="f0a40-117">The custom functions runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes.</span></span> <span data-ttu-id="f0a40-118">它们都称为“对话框 API”，但在自定义函数运行时中使用 `OfficeRuntime.Dialog` 对用户进行身份验证。</span><span class="sxs-lookup"><span data-stu-id="f0a40-118">They're both referred to as the "Dialog API", but use `OfficeRuntime.Dialog` to authenticate users in the custom functions runtime.</span></span>

<span data-ttu-id="f0a40-119">有关如何使用 `Dialog` 对象的信息，请参阅[自定义函数对话框](/office/dev/add-ins/excel/custom-functions-dialog)。</span><span class="sxs-lookup"><span data-stu-id="f0a40-119">For information on how to use the `Dialog` object, see [Custom Functions dialog](/office/dev/add-ins/excel/custom-functions-dialog).</span></span>

<span data-ttu-id="f0a40-120">在构想整个身份验证过程时，将加载项的任务窗格和 UI 元素以及加载项的自定义函数部分视为可以通过 `OfficeRuntime.storage` 进行相互通信的单独实体，这样做可能对你有所帮助。</span><span class="sxs-lookup"><span data-stu-id="f0a40-120">When envisioning the entire authentication process as a whole, it might be helpful to think of the task pane and UI elements of your add-in and the custom functions part of your add-in as separate entities which can communicate with each other through `OfficeRuntime.storage`.</span></span>

<span data-ttu-id="f0a40-121">下图概述了此基本过程。</span><span class="sxs-lookup"><span data-stu-id="f0a40-121">The following diagram outlines this basic process.</span></span> <span data-ttu-id="f0a40-122">请注意，虚线指示虽然它们执行单独的操作，但自定义函数和加载项的任务窗格都是整个加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="f0a40-122">Note that the dotted line indicates that while they perform separate actions, custom functions and your add-in's task pane are both part of your add-in as a whole.</span></span>

1. <span data-ttu-id="f0a40-123">你可以从 Excel 工作簿中的单元格发出自定义函数调用。</span><span class="sxs-lookup"><span data-stu-id="f0a40-123">You issue a custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="f0a40-124">自定义函数使用 `Dialog` 将你的用户凭据传递给网站。</span><span class="sxs-lookup"><span data-stu-id="f0a40-124">The custom function uses `Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="f0a40-125">该网站随后会将访问令牌返回给自定义函数。</span><span class="sxs-lookup"><span data-stu-id="f0a40-125">This website then returns an access token to the custom function.</span></span>
4. <span data-ttu-id="f0a40-126">然后，自定义函数会将此访问令牌存储在 `storage` 中。</span><span class="sxs-lookup"><span data-stu-id="f0a40-126">Your custom function then sets this access token to the `storage`.</span></span>
5. <span data-ttu-id="f0a40-127">加载项的任务窗格将从 `storage` 访问该令牌。</span><span class="sxs-lookup"><span data-stu-id="f0a40-127">Your add-in's task pane accesses the token from `storage`.</span></span>

<span data-ttu-id="f0a40-128">![使用对话框 API 获取访问令牌并通过 OfficeRuntime.storage API 与任务窗格共享令牌的自定义函数关系图。](../images/authentication-diagram.png "身份验证关系图。")</span><span class="sxs-lookup"><span data-stu-id="f0a40-128">![Diagram of custom function using dialog API to get access token, and then share token with task pane through the OfficeRuntime.storage API.](../images/authentication-diagram.png "Authentication diagram.")</span></span>

## <a name="storing-the-token"></a><span data-ttu-id="f0a40-129">存储令牌</span><span class="sxs-lookup"><span data-stu-id="f0a40-129">Storing the token</span></span>

<span data-ttu-id="f0a40-130">以下示例来自[在自定义函数中使用 OfficeRuntime.storage](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) 代码示例。</span><span class="sxs-lookup"><span data-stu-id="f0a40-130">The following examples are from the [Using OfficeRuntime.storage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample.</span></span> <span data-ttu-id="f0a40-131">有关在自定义函数与任务窗格之间共享数据的完整示例，请参阅此代码示例。</span><span class="sxs-lookup"><span data-stu-id="f0a40-131">Refer to this code sample for a complete example of sharing data between custom functions and the task pane.</span></span>

<span data-ttu-id="f0a40-132">如果使用自定义函数进行身份验证，则它会收到访问令牌，并且需要将其存储在 `storage` 中。</span><span class="sxs-lookup"><span data-stu-id="f0a40-132">If the custom function authenticates, then it receives the access token and will need to store it in `storage`.</span></span> <span data-ttu-id="f0a40-133">以下代码示例演示如何调用 `storage.setItem` 方法来存储值。</span><span class="sxs-lookup"><span data-stu-id="f0a40-133">The following code sample shows how to call the `storage.setItem` method to store a value.</span></span> <span data-ttu-id="f0a40-134">`storeValue` 函数是一个自定义函数，例如用于存储用户的值。</span><span class="sxs-lookup"><span data-stu-id="f0a40-134">The `storeValue` function is a custom function that for example purposes stores a value from the user.</span></span> <span data-ttu-id="f0a40-135">你可以对其进行修改以存储所需的任何令牌值。</span><span class="sxs-lookup"><span data-stu-id="f0a40-135">You can modify this to store any token value you need.</span></span>

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

<span data-ttu-id="f0a40-136">当任务窗格需要访问令牌时，它可以从 `storage` 检索令牌。</span><span class="sxs-lookup"><span data-stu-id="f0a40-136">When the task pane needs the access token, it can retrieve the token from `storage`.</span></span> <span data-ttu-id="f0a40-137">以下代码示例演示如何使用 `storage.getItem` 方法来检索令牌。</span><span class="sxs-lookup"><span data-stu-id="f0a40-137">The following code sample shows how to use the `storage.getItem` method to retrieve the token.</span></span>

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

## <a name="general-guidance"></a><span data-ttu-id="f0a40-138">一般指导</span><span class="sxs-lookup"><span data-stu-id="f0a40-138">General Guidance</span></span>

<span data-ttu-id="f0a40-139">Office 加载项基于 Web，你可以使用任何 Web 身份验证技术。</span><span class="sxs-lookup"><span data-stu-id="f0a40-139">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="f0a40-140">使用自定义函数实施自己的身份验证时，不必遵循特定的模式或方法。</span><span class="sxs-lookup"><span data-stu-id="f0a40-140">There is no particular pattern or method you must follow to implement your own authentication with custom functions.</span></span> <span data-ttu-id="f0a40-141">你可能希望查阅有关各种身份验证模式的文档，请先参阅[这篇关于通过外部服务进行授权的文章](/office/dev/add-ins/develop/auth-external-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="f0a40-141">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](/office/dev/add-ins/develop/auth-external-add-ins).</span></span>  

<span data-ttu-id="f0a40-142">在开发自定义函数时，避免使用以下位置存储数据：</span><span class="sxs-lookup"><span data-stu-id="f0a40-142">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="f0a40-143">`localStorage`：自定义函数无权访问全局 `window` 对象，因此无法访问 `localStorage` 中存储的数据。</span><span class="sxs-lookup"><span data-stu-id="f0a40-143">`localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data stored in `localStorage`.</span></span>
- <span data-ttu-id="f0a40-144">`Office.context.document.settings`：此位置不安全，使用加载项的任何人员都可以提取相关信息。</span><span class="sxs-lookup"><span data-stu-id="f0a40-144">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the add-in.</span></span>

## <a name="next-steps"></a><span data-ttu-id="f0a40-145">后续步骤</span><span class="sxs-lookup"><span data-stu-id="f0a40-145">Next steps</span></span>
<span data-ttu-id="f0a40-146">了解[自定义函数的对话框 API](custom-functions-dialog.md)。</span><span class="sxs-lookup"><span data-stu-id="f0a40-146">Learn about the [dialog API for custom functions](custom-functions-dialog.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f0a40-147">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f0a40-147">See also</span></span>

* [<span data-ttu-id="f0a40-148">自定义函数体系结构</span><span class="sxs-lookup"><span data-stu-id="f0a40-148">Custom functions architecture</span></span>](custom-functions-architecture.md)
* [<span data-ttu-id="f0a40-149">使用自定义函数接收和处理数据</span><span class="sxs-lookup"><span data-stu-id="f0a40-149">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="f0a40-150">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="f0a40-150">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="f0a40-151">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="f0a40-151">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)

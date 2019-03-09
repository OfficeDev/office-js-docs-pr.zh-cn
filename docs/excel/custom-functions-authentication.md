---
ms.date: 03/06/2019
description: 在 Excel 中使用自定义函数对用户进行身份验证。
title: 自定义函数的身份验证
ms.openlocfilehash: 4358d9f570ef8b31db98b1886c01ff4a89a6b1be
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512851"
---
# <a name="authentication"></a><span data-ttu-id="75553-103">身份验证</span><span class="sxs-lookup"><span data-stu-id="75553-103">Authentication</span></span>

<span data-ttu-id="75553-104">在某些情况下, 自定义函数将需要对用户进行身份验证, 以便访问受保护的资源。</span><span class="sxs-lookup"><span data-stu-id="75553-104">In some scenarios your custom function will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="75553-105">虽然自定义函数不需要特定的身份验证方法, 但您应注意, 自定义函数在单独的运行时中从任务窗格和外接程序的其他 UI 元素运行。</span><span class="sxs-lookup"><span data-stu-id="75553-105">While custom functions don't require a specific method of authentication, you should be aware that custom functions run in a separate runtime from the task pane and other UI elements of your add-in.</span></span> <span data-ttu-id="75553-106">因此, 您需要使用`AsyncStorage`对象和对话框 API 在两个运行时之间来回传递数据。</span><span class="sxs-lookup"><span data-stu-id="75553-106">Because of this, you'll need to pass data back and forth between the two runtimes using the `AsyncStorage` object and the Dialog API.</span></span>
  
## <a name="asyncstorage-object"></a><span data-ttu-id="75553-107">到 asyncstorage 对象</span><span class="sxs-lookup"><span data-stu-id="75553-107">AsyncStorage object</span></span>

<span data-ttu-id="75553-108">自定义函数运行时在全局`localStorage`窗口中没有可用的对象, 您通常可能会在其中存储数据。</span><span class="sxs-lookup"><span data-stu-id="75553-108">The custom functions runtime doesn't have a `localStorage` object available on the global window, where you might typically store data.</span></span> <span data-ttu-id="75553-109">相反, 您应该使用[OfficeRuntime](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage)来设置和获取数据, 从而在自定义函数和任务窗格之间共享数据。</span><span class="sxs-lookup"><span data-stu-id="75553-109">Instead, you should share data between custom functions and task panes by using [OfficeRuntime.AsyncStorage](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage) to set and get data.</span></span>

<span data-ttu-id="75553-110">此外, 还提供了使用`AsyncStorage`的好处;它使用安全沙盒环境, 以便其他外接程序无法访问您的数据。</span><span class="sxs-lookup"><span data-stu-id="75553-110">Additionally, there is a benefit to using `AsyncStorage`; it uses a secure sandbox environment so that your data cannot be accessed by other add-ins.</span></span>

### <a name="suggested-usage"></a><span data-ttu-id="75553-111">建议使用</span><span class="sxs-lookup"><span data-stu-id="75553-111">Suggested usage</span></span>

<span data-ttu-id="75553-112">当您需要从任务窗格或自定义函数进行身份验证时, 请`AsyncStorage`检查是否已获取访问令牌。</span><span class="sxs-lookup"><span data-stu-id="75553-112">When you need to authenticate either from the task pane or a custom function, check `AsyncStorage` to see if the access token was already acquired.</span></span> <span data-ttu-id="75553-113">如果不是, 请使用对话框 API 对用户进行身份验证、检索访问令牌, 然后将令牌存储在`AsyncStorage`中以备将来使用。</span><span class="sxs-lookup"><span data-stu-id="75553-113">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `AsyncStorage` for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="75553-114">对话框 API</span><span class="sxs-lookup"><span data-stu-id="75553-114">Dialog API</span></span>

<span data-ttu-id="75553-115">如果令牌不存在, 则应使用对话框 API 要求用户登录。</span><span class="sxs-lookup"><span data-stu-id="75553-115">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="75553-116">用户输入凭据后, 生成的访问令牌可以存储在中`AsyncStorage`。</span><span class="sxs-lookup"><span data-stu-id="75553-116">After a user enters their credentials, the resulting access token can be stored in `AsyncStorage`.</span></span>

> [!NOTE]
> <span data-ttu-id="75553-117">自定义函数运行时使用与任务窗格使用的浏览器引擎运行时中的 dialog 对象略有不同的 dialog 对象。</span><span class="sxs-lookup"><span data-stu-id="75553-117">The custom functions runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes.</span></span> <span data-ttu-id="75553-118">它们都称为 "对话框 API", 但用于`Officeruntime.Dialog`在自定义函数运行时中对用户进行身份验证。</span><span class="sxs-lookup"><span data-stu-id="75553-118">They're both referred to as the "Dialog API", but use `Officeruntime.Dialog` to authenticate users in the custom functions runtime.</span></span>

<span data-ttu-id="75553-119">有关如何使用的`OfficeRuntime.Dialog`信息, 请参阅[Custom 函数运行时](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box)。</span><span class="sxs-lookup"><span data-stu-id="75553-119">For information on how to use the `OfficeRuntime.Dialog`, see [Custom Functions runtime](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box).</span></span>

<span data-ttu-id="75553-120">在整体上构思整个身份验证过程时, 将加载项的任务窗格和 UI 元素以及外接程序的自定义函数部分视为可通过`AsyncStorage`相互通信的单独实体可能会有所帮助。</span><span class="sxs-lookup"><span data-stu-id="75553-120">When envisioning the entire authentication process as a whole, it might be helpful to think of the task pane and UI elements of your add-in and the custom functions part of your add-in as separate entities which can communicate with each other through `AsyncStorage`.</span></span>

<span data-ttu-id="75553-121">下图概述了此基本过程。</span><span class="sxs-lookup"><span data-stu-id="75553-121">The following diagram outlines this basic process.</span></span> <span data-ttu-id="75553-122">请注意, 点线表示在执行单独的操作时, 自定义函数和外接程序的任务窗格都是外接程序的整体。</span><span class="sxs-lookup"><span data-stu-id="75553-122">Note that the dotted line indicates that while they perform separate actions, custom functions and your add-in's task pane are both part of your add-in as a whole.</span></span>

1. <span data-ttu-id="75553-123">您从 Excel 工作簿中的单元格发出自定义函数调用。</span><span class="sxs-lookup"><span data-stu-id="75553-123">You issue a custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="75553-124">自定义函数`Officeruntime.Dialog`用于将您的用户凭据传递到网站。</span><span class="sxs-lookup"><span data-stu-id="75553-124">The custom function uses `Officeruntime.Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="75553-125">然后, 此网站将向自定义函数返回访问令牌。</span><span class="sxs-lookup"><span data-stu-id="75553-125">This website then returns an access token to the custom function.</span></span>
4. <span data-ttu-id="75553-126">然后, 您的`AsyncStorage`自定义函数会将此访问令牌设置为。</span><span class="sxs-lookup"><span data-stu-id="75553-126">Your custom function then sets this access token to the `AsyncStorage`.</span></span>
5. <span data-ttu-id="75553-127">外接程序的任务窗格从`AsyncStorage`访问令牌。</span><span class="sxs-lookup"><span data-stu-id="75553-127">Your add-in's task pane accesses the token from `AsyncStorage`.</span></span>

<span data-ttu-id="75553-128">![协同工作的自定义函数、OfficeRuntime 和任务窗格的关系图。](../images/Authdiagram.png "身份验证图。")</span><span class="sxs-lookup"><span data-stu-id="75553-128">![Diagram of custom functions, OfficeRuntime, and task panes working together.](../images/Authdiagram.png "Authentication diagram.")</span></span>

## <a name="storing-the-token"></a><span data-ttu-id="75553-129">存储令牌</span><span class="sxs-lookup"><span data-stu-id="75553-129">Storing the token</span></span>

<span data-ttu-id="75553-130">下面的示例来自[自定义函数代码示例中的 Using 到 asyncstorage](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) 。</span><span class="sxs-lookup"><span data-stu-id="75553-130">The following examples are from the [Using AsyncStorage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample.</span></span> <span data-ttu-id="75553-131">有关在自定义函数和任务窗格之间共享数据的完整示例, 请参阅此代码示例。</span><span class="sxs-lookup"><span data-stu-id="75553-131">Refer to this code sample for a complete example of sharing data between custom functions and the task pane.</span></span>

<span data-ttu-id="75553-132">如果自定义函数进行身份验证, 则它将接收访问令牌, 并需要将其`AsyncStorage`存储在中。</span><span class="sxs-lookup"><span data-stu-id="75553-132">If the custom function authenticates, then it receives the access token and will need to store it in `AsyncStorage`.</span></span> <span data-ttu-id="75553-133">下面的代码示例演示如何调用`AsyncStorage.setItem`方法来存储值。</span><span class="sxs-lookup"><span data-stu-id="75553-133">The following code sample shows how to call the `AsyncStorage.setItem` method to store a value.</span></span> <span data-ttu-id="75553-134">`StoreValue`函数是一个自定义函数, 例如, 用于存储来自用户的值。</span><span class="sxs-lookup"><span data-stu-id="75553-134">The `StoreValue` function is a custom function that for example purposes stores a value from the user.</span></span> <span data-ttu-id="75553-135">您可以对此进行修改, 以存储所需的任何标记值。</span><span class="sxs-lookup"><span data-stu-id="75553-135">You can modify this to store any token value you need.</span></span>

```javascript
function StoreValue(key, value) {
  return OfficeRuntime.AsyncStorage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to AsyncStorage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to AsyncStorage. " + error;
  });
}
```

<span data-ttu-id="75553-136">当任务窗格需要访问令牌时, 它可以从`AsyncStorage`检索令牌。</span><span class="sxs-lookup"><span data-stu-id="75553-136">When the task pane needs the access token, it can retrieve the token from `AsyncStorage`.</span></span> <span data-ttu-id="75553-137">下面的代码示例演示如何使用`AsyncStorage.getItem`方法检索令牌。</span><span class="sxs-lookup"><span data-stu-id="75553-137">The following code sample shows how to use the `AsyncStorage.getItem` method to retrieve the token.</span></span>

```javascript
function ReceiveTokenFromCustomFunction() {
   var key = "token";
   var tokenSendStatus = document.getElementById('tokenSendStatus');
   OfficeRuntime.AsyncStorage.getItem(key).then(function (result) {
      tokenSendStatus.value = "Success: Item with key '" + key + "' read from AsyncStorage.";
      document.getElementById('tokenTextBox2').value = result;
   }, function (error) {
      tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from AsyncStorage. " + error;
   });
}
```

## <a name="general-guidance"></a><span data-ttu-id="75553-138">一般指南</span><span class="sxs-lookup"><span data-stu-id="75553-138">General guidance</span></span>

<span data-ttu-id="75553-139">Office 外接程序是基于 web 的, 您可以使用任何 web 身份验证技术。</span><span class="sxs-lookup"><span data-stu-id="75553-139">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="75553-140">使用自定义函数实现自己的身份验证时, 必须遵循任何特定的模式或方法。</span><span class="sxs-lookup"><span data-stu-id="75553-140">There is no particular pattern or method you must follow to implement your own authentication with custom functions.</span></span> <span data-ttu-id="75553-141">您可能希望参考有关各种身份验证模式的文档, 从本文开始,[了解如何通过外部服务进行授权](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js)。</span><span class="sxs-lookup"><span data-stu-id="75553-141">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).</span></span>  

<span data-ttu-id="75553-142">在开发自定义函数时, 应避免使用以下位置来存储数据:</span><span class="sxs-lookup"><span data-stu-id="75553-142">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="75553-143">`localStorage`: 自定义函数不具有对全局`window`对象的访问权限, 因此无法访问存储在中`localStorage`的数据。</span><span class="sxs-lookup"><span data-stu-id="75553-143">`localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data     stored in `localStorage`.</span></span>
- <span data-ttu-id="75553-144">`Office.context.document.settings`: 此位置不安全, 使用外接程序的任何人都可以提取信息。</span><span class="sxs-lookup"><span data-stu-id="75553-144">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the     add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="75553-145">另请参阅</span><span class="sxs-lookup"><span data-stu-id="75553-145">See also</span></span>

* [<span data-ttu-id="75553-146">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="75553-146">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="75553-147">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="75553-147">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="75553-148">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="75553-148">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="75553-149">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="75553-149">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)

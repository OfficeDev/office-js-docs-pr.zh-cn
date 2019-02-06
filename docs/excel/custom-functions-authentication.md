---
ms.date: 1/29/2019
description: 在 Excel 中使用自定义函数的用户进行身份验证。
title: 身份验证的自定义的函数
ms.openlocfilehash: 0e42dbc93cb545660a8dbaae5bdb48724f3b7376
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/05/2019
ms.locfileid: "29745403"
---
# <a name="authentication"></a><span data-ttu-id="52b75-103">身份验证</span><span class="sxs-lookup"><span data-stu-id="52b75-103">Authentication</span></span>

<span data-ttu-id="52b75-104">在某些情况下，您自定义的函数将需要对用户进行身份验证才能访问受保护资源。</span><span class="sxs-lookup"><span data-stu-id="52b75-104">In some scenarios your custom function will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="52b75-105">自定义函数不需要特定的身份验证方法时, 应注意的自定义函数在单独的运行时从运行任务窗格和加载项的其他用户界面元素。</span><span class="sxs-lookup"><span data-stu-id="52b75-105">While custom functions doesn't require a specific method of authentication, you should be aware that custom functions runs in a separate runtime from the task pane and other UI elements of your add-in.</span></span> <span data-ttu-id="52b75-106">因此，您需要使用两个运行时之间来回传递数据`AsyncStorage`对象和对话框 API。</span><span class="sxs-lookup"><span data-stu-id="52b75-106">Because of this, you'll need to pass data back and forth between the two runtimes using the `AsyncStorage` object and the Dialog API.</span></span>
  
## <a name="asyncstorage-object"></a><span data-ttu-id="52b75-107">AsyncStorage 对象</span><span class="sxs-lookup"><span data-stu-id="52b75-107">AsyncStorage object</span></span>

<span data-ttu-id="52b75-108">自定义函数运行时没有`localStorage`上全局窗口，其中可能通常存储数据的可用对象。</span><span class="sxs-lookup"><span data-stu-id="52b75-108">The custom functions runtime doesn't have a `localStorage` object available on the global window, where you might typically store data.</span></span> <span data-ttu-id="52b75-109">相反，您应之间共享数据自定义函数和任务窗格中，通过使用[OfficeRuntime.AsyncStorage](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage)设置和获取数据。</span><span class="sxs-lookup"><span data-stu-id="52b75-109">Instead, you should share data between custom functions and task panes, by using [OfficeRuntime.AsyncStorage](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage) to set and get data.</span></span> 

<span data-ttu-id="52b75-110">此外，还会使用好处`AsyncStorage`;它使用安全沙盒环境，以便其他加载项无法访问您的数据。</span><span class="sxs-lookup"><span data-stu-id="52b75-110">Additionally, there is a benefit to using `AsyncStorage`; it uses a secure sandbox environment so that your data cannot be accessed by other add-ins.</span></span>  

### <a name="suggested-usage"></a><span data-ttu-id="52b75-111">建议使用情况</span><span class="sxs-lookup"><span data-stu-id="52b75-111">Suggested usage</span></span>

<span data-ttu-id="52b75-112">当您需要进行身份验证从任务窗格或自定义的函数时，检查 AsyncStorage 以查看是否已被收购访问令牌。</span><span class="sxs-lookup"><span data-stu-id="52b75-112">When you need to authenticate either from the task pane or a custom function, check AsyncStorage to see if the access token was already acquired.</span></span> <span data-ttu-id="52b75-113">如果没有，请使用对话框 API 来验证用户身份和检索访问令牌，然后将该令牌存储在 AsyncStorage 以供将来使用。</span><span class="sxs-lookup"><span data-stu-id="52b75-113">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in AsyncStorage for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="52b75-114">对话框 API</span><span class="sxs-lookup"><span data-stu-id="52b75-114">Dialog API</span></span>

<span data-ttu-id="52b75-115">如果令牌不存在，您应使用对话框 API 要求用户登录。</span><span class="sxs-lookup"><span data-stu-id="52b75-115">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="52b75-116">生成访问令牌用户输入凭据之后，可以存储在`AsyncStorage`。</span><span class="sxs-lookup"><span data-stu-id="52b75-116">After a user enters their credentials, the resulting access token can be stored in `AsyncStorage`.</span></span>

> [!NOTE]
> <span data-ttu-id="52b75-117">自定义函数运行时使用此对话框对象中运行时使用的任务窗格略有不同 Dialog 对象。</span><span class="sxs-lookup"><span data-stu-id="52b75-117">The custom functions runtime uses a Dialog object that is slightly different from the Dialog object in the runtime used by task panes.</span></span> <span data-ttu-id="52b75-118">它们同时称为"对话框 API"，但使用`Officeruntime.Dialog`中的自定义函数的运行时的用户进行身份验证。</span><span class="sxs-lookup"><span data-stu-id="52b75-118">They're both referred to as the "Dialog API", but use `Officeruntime.Dialog` to authenticate users in the custom functions runtime.</span></span>

<span data-ttu-id="52b75-119">有关如何使用`OfficeRuntime.Dialog`，请参阅[自定义函数的运行时](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box)。</span><span class="sxs-lookup"><span data-stu-id="52b75-119">For information on how to use the `OfficeRuntime.Dialog`, see [Custom Functions runtime](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box).</span></span>

<span data-ttu-id="52b75-120">时构想作为一个整体整个身份验证过程，它可能需要考虑的任务窗格和加载项的 UI 元素和自定义为单独的实体，其中可以与通过每个其他通信功能的加载项部分`AsyncStorage`。</span><span class="sxs-lookup"><span data-stu-id="52b75-120">When envisioning the entire authentication process as a whole, it might be helpful to think of the task pane and UI elements of your add-in and the custom functions portions of your add-in as separate entities which can communicate with each other through `AsyncStorage`.</span></span>

<span data-ttu-id="52b75-121">下图概述了此基本过程。</span><span class="sxs-lookup"><span data-stu-id="52b75-121">The following diagram outlines this basic process.</span></span> <span data-ttu-id="52b75-122">请注意，虚线表示他们执行单独操作，自定义的函数和外接程序的任务窗格的加载项作为一个整体两个部分。</span><span class="sxs-lookup"><span data-stu-id="52b75-122">Note that the dotted line indicates that while they perform separate actions, custom functions and your add-in's task pane are both parts of your add-in as a whole.</span></span>

1. <span data-ttu-id="52b75-123">问题自定义的函数调用从 Excel 工作簿中的单元格。</span><span class="sxs-lookup"><span data-stu-id="52b75-123">You issue a custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="52b75-124">自定义的函数使用`Officeruntime.Dialog`可以将您的用户凭据传递到网站。</span><span class="sxs-lookup"><span data-stu-id="52b75-124">The custom function uses `Officeruntime.Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="52b75-125">然后，该网站将访问令牌返回到自定义的函数。</span><span class="sxs-lookup"><span data-stu-id="52b75-125">This website then returns an access token to the custom function.</span></span>
4. <span data-ttu-id="52b75-126">然后，您自定义的函数将此访问令牌设置为`AsyncStorage`。</span><span class="sxs-lookup"><span data-stu-id="52b75-126">Your custom function then sets this access token to the `AsyncStorage`.</span></span>
5. <span data-ttu-id="52b75-127">外接程序的任务窗格访问来自令牌`AsyncStorage`。</span><span class="sxs-lookup"><span data-stu-id="52b75-127">Your add-in's task pane accesses the token from `AsyncStorage`.</span></span>

<span data-ttu-id="52b75-128">![的自定义的函数、 OfficeRuntime 和协作的任务窗格的图表。](../images/Authdiagram.png "身份验证关系图。")</span><span class="sxs-lookup"><span data-stu-id="52b75-128">![Diagram of custom functions, OfficeRuntime, and task panes working together.](../images/Authdiagram.png "Authentication diagram.")</span></span>

## <a name="general-guidance"></a><span data-ttu-id="52b75-129">一般指导</span><span class="sxs-lookup"><span data-stu-id="52b75-129">General guidance</span></span>

<span data-ttu-id="52b75-130">Office 加载项是基于 web 的您可以使用任何 web 身份验证方法。</span><span class="sxs-lookup"><span data-stu-id="52b75-130">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="52b75-131">没有特定模式或方法必须遵循实现自己的身份验证与自定义函数。</span><span class="sxs-lookup"><span data-stu-id="52b75-131">There is no particular pattern or method you must follow to implement your own authentication with custom functions.</span></span> <span data-ttu-id="52b75-132">您可能希望查阅有关各种身份验证模式，文档开头[有关授权通过外部服务这篇文章](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js)。</span><span class="sxs-lookup"><span data-stu-id="52b75-132">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).</span></span>  

<span data-ttu-id="52b75-133">避免使用的以下位置来开发自定义的函数时存储数据：</span><span class="sxs-lookup"><span data-stu-id="52b75-133">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="52b75-134">`localStorage`： 自定义函数不具有对全局访问`window`对象，并因此均没有访问权数据存储在`localStorage`。</span><span class="sxs-lookup"><span data-stu-id="52b75-134">`localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data     stored in `localStorage`.</span></span>
- <span data-ttu-id="52b75-135">`Office.context.document.settings`： 此位置不安全，通过使用外接程序的任何人都可以提取信息。</span><span class="sxs-lookup"><span data-stu-id="52b75-135">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the     add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="52b75-136">另请参阅</span><span class="sxs-lookup"><span data-stu-id="52b75-136">See also</span></span>

* [<span data-ttu-id="52b75-137">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="52b75-137">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="52b75-138">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="52b75-138">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="52b75-139">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="52b75-139">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="52b75-140">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="52b75-140">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)

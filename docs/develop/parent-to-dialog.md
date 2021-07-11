---
title: 将邮件从主机页传递到对话框的替代方法
description: 了解在 messageChild 方法不受支持时使用的解决方法。
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: 8da6bc3e1231bc6296a16fa153dc0e4ba1bd102b
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349775"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a><span data-ttu-id="86119-103">将邮件从主机页传递到对话框的替代方法</span><span class="sxs-lookup"><span data-stu-id="86119-103">Alternative ways of passing messages to a dialog box from its host page</span></span>

<span data-ttu-id="86119-104">将数据和消息从父页面传递到子对话框的建议方法是使用 方法，如在 Office 加载项中使用 Office 对话框 `messageChild` [API 中所述](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)。如果加载项运行在不支持[DialogApi 1.2](../reference/requirement-sets/dialog-api-requirement-sets.md)要求集的平台或主机上，有两种其他方法可以将信息传递到对话框：</span><span class="sxs-lookup"><span data-stu-id="86119-104">The recommended way to pass data and messages from a parent page to a child dialog box is with the `messageChild` method as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box). If your add-in is running on a platform or host that does not support the [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md), there are two other ways that you can pass information to the dialog box:</span></span>

- <span data-ttu-id="86119-105">向传递给 `displayDialogAsync` 的 URL 添加查询参数。</span><span class="sxs-lookup"><span data-stu-id="86119-105">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="86119-106">将信息存储在主机窗口和对话框都可访问的位置。</span><span class="sxs-lookup"><span data-stu-id="86119-106">Store the information somewhere that is accessible to both the host window and dialog box.</span></span> <span data-ttu-id="86119-107">这两个窗口不共享 [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage))  (的常见会话存储，但如果它们具有相同的域 *(包括* 端口号，如果有) ，则它们共享一个公共 [本地 存储](https://www.w3schools.com/html/html5_webstorage.asp)。\*</span><span class="sxs-lookup"><span data-stu-id="86119-107">The two windows do not share a common session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property), but *if they have the same domain* (including port number, if any), they share a common [Local Storage](https://www.w3schools.com/html/html5_webstorage.asp).\*</span></span>


> [!NOTE]
> <span data-ttu-id="86119-108">\*有一个 bug 将影响你的令牌处理策略。</span><span class="sxs-lookup"><span data-stu-id="86119-108">\* There is a bug that will effect your strategy for token handling.</span></span> <span data-ttu-id="86119-109">如果加载项正使用 Safari 或 Microsoft 浏览器在 **Office 网页版** 上运行，则对话框和任务窗格不共享同一本地存储，因此该存储无法用于在它们之间通信。</span><span class="sxs-lookup"><span data-stu-id="86119-109">If the add-in is running in **Office on the web** in either the Safari or Edge browser, the dialog box and task pane do not share the same Local Storage, so it cannot be used to communicate between them.</span></span>

## <a name="use-local-storage"></a><span data-ttu-id="86119-110">使用本地存储</span><span class="sxs-lookup"><span data-stu-id="86119-110">Use local storage</span></span>

<span data-ttu-id="86119-111">若要使用本地存储，在调用前调用主机页中的 对象的 方法 `setItem` `window.localStorage` `displayDialogAsync` ，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="86119-111">To use local storage, call the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example.</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="86119-112">对话框中的代码会根据需要读取项目，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="86119-112">Code in the dialog box reads the item when it's needed, as in the following example.</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a><span data-ttu-id="86119-113">使用查询参数</span><span class="sxs-lookup"><span data-stu-id="86119-113">Use query parameters</span></span>

<span data-ttu-id="86119-114">下面的示例展示了如何使用查询参数传递数据。</span><span class="sxs-lookup"><span data-stu-id="86119-114">The following example shows how to pass data with a query parameter.</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="86119-115">有关使用此技术的示例，请参阅[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)（在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表）。</span><span class="sxs-lookup"><span data-stu-id="86119-115">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="86119-116">对话框中的代码可以分析 URL，并读取参数值。</span><span class="sxs-lookup"><span data-stu-id="86119-116">Code in your dialog box can parse the URL and read the parameter value.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="86119-117">Office 会自动向传递给 `_host_info` 的 URL 添加查询参数 `displayDialogAsync`。</span><span class="sxs-lookup"><span data-stu-id="86119-117">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`.</span></span> <span data-ttu-id="86119-118">（附加在自定义查询参数（若有）后面，不会附加到对话框导航到的任何后续 URL。</span><span class="sxs-lookup"><span data-stu-id="86119-118">(It is appended after your custom query parameters, if any.</span></span> <span data-ttu-id="86119-119">）Microsoft 可能会更改此值的内容，或者将来会将其全部删除，因此代码不得读取此值。</span><span class="sxs-lookup"><span data-stu-id="86119-119">It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it.</span></span> <span data-ttu-id="86119-120">相同的值将添加到对话框的会话存储 ([Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)) 。</span><span class="sxs-lookup"><span data-stu-id="86119-120">The same value is added to the dialog box's session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property).</span></span> <span data-ttu-id="86119-121">同样，*代码不得对此值执行读取和写入操作*。</span><span class="sxs-lookup"><span data-stu-id="86119-121">Again, *your code should neither read nor write to this value*.</span></span>

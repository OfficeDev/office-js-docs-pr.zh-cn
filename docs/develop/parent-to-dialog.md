---
title: 将邮件从其主页面传递到对话框的替代方法
description: 了解在 messageChild 方法不受支持时要使用的解决方法。
ms.date: 08/20/2020
localization_priority: Normal
ms.openlocfilehash: b516896d28979f439f3065f9ff036ff21c2c0997
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293175"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a><span data-ttu-id="e9481-103">将邮件从其主页面传递到对话框的替代方法</span><span class="sxs-lookup"><span data-stu-id="e9481-103">Alternative ways of passing messages to a dialog box from its host page</span></span>

<span data-ttu-id="e9481-104">将来自父页面的数据和邮件传递到子对话框的建议方法是 `messageChild` 使用方法，如在 [office 外接程序中使用 OFFICE 对话框 API](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)中所述。如果外接程序在不支持 [DialogApi 1.2 要求集](../reference/requirement-sets/dialog-api-requirement-sets.md)的平台或主机上运行，则可以通过两种其他方式将信息传递到该对话框：</span><span class="sxs-lookup"><span data-stu-id="e9481-104">The recommended way to pass data and messages from a parent page to a child dialog box is with the `messageChild` method as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box). If your add-in is running on a platform or host that does not support the [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md), there are two other ways that you can pass information to the dialog box:</span></span>

- <span data-ttu-id="e9481-105">向传递给 `displayDialogAsync` 的 URL 添加查询参数。</span><span class="sxs-lookup"><span data-stu-id="e9481-105">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="e9481-106">将信息存储在主机窗口和对话框都可访问的位置。</span><span class="sxs-lookup"><span data-stu-id="e9481-106">Store the information somewhere that is accessible to both the host window and dialog box.</span></span> <span data-ttu-id="e9481-107">这两个窗口不共享通用会话存储，但*如果它们具有相同的域*（包括端口号，若有），则共享通用[本地存储](https://www.w3schools.com/html/html5_webstorage.asp)。\*</span><span class="sxs-lookup"><span data-stu-id="e9481-107">The two windows do not share a common session storage, but *if they have the same domain* (including port number, if any), they share a common [Local Storage](https://www.w3schools.com/html/html5_webstorage.asp).\*</span></span>


> [!NOTE]
> <span data-ttu-id="e9481-108">\*有一个 bug 将影响你的令牌处理策略。</span><span class="sxs-lookup"><span data-stu-id="e9481-108">\* There is a bug that will effect your strategy for token handling.</span></span> <span data-ttu-id="e9481-109">如果加载项正使用 Safari 或 Microsoft 浏览器在 **Office 网页版**上运行，则对话框和任务窗格不共享同一本地存储，因此该存储无法用于在它们之间通信。</span><span class="sxs-lookup"><span data-stu-id="e9481-109">If the add-in is running in **Office on the web** in either the Safari or Edge browser, the dialog box and task pane do not share the same Local Storage, so it cannot be used to communicate between them.</span></span>

## <a name="use-local-storage"></a><span data-ttu-id="e9481-110">使用本地存储</span><span class="sxs-lookup"><span data-stu-id="e9481-110">Use local storage</span></span>

<span data-ttu-id="e9481-111">若要使用本地存储，请 `setItem` `window.localStorage` 先在主机页中调用对象的方法，然后再 `displayDialogAsync` 调用，如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="e9481-111">To use local storage, call the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="e9481-112">对话框框中的代码会在需要时读取项，如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="e9481-112">Code in the dialog box reads the item when it's needed, as in the following example:</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a><span data-ttu-id="e9481-113">使用查询参数</span><span class="sxs-lookup"><span data-stu-id="e9481-113">Use query parameters</span></span>

<span data-ttu-id="e9481-114">下面的示例展示了如何使用查询参数传递数据：</span><span class="sxs-lookup"><span data-stu-id="e9481-114">The following example shows how to pass data with a query parameter:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="e9481-115">有关使用此技术的示例，请参阅[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)（在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表）。</span><span class="sxs-lookup"><span data-stu-id="e9481-115">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="e9481-116">对话框中的代码可以分析 URL，并读取参数值。</span><span class="sxs-lookup"><span data-stu-id="e9481-116">Code in your dialog box can parse the URL and read the parameter value.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e9481-p103">Office 会自动向传递给 `displayDialogAsync` 的 URL 添加查询参数 `_host_info`。（附加在自定义查询参数（若有）后面，不会附加到对话框导航到的任何后续 URL。）Microsoft 可能会更改此值的内容，或者将来会将其全部删除，因此代码不得读取此值。相同的值会被添加到对话框的会话存储中。同样，*代码不得对此值执行读取和写入操作*。</span><span class="sxs-lookup"><span data-stu-id="e9481-p103">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.</span></span>

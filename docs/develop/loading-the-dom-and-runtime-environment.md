---
title: 加载 DOM 和运行时环境
description: ''
ms.date: 01/09/2019
localization_priority: Priority
ms.openlocfilehash: c569825ae73d32259fc1554aa8233461bbbe9813
ms.sourcegitcommit: 70ef38a290c18a1d1a380fd02b263470207a5dc6
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/15/2019
ms.locfileid: "30052747"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="10b03-102">加载 DOM 和运行时环境</span><span class="sxs-lookup"><span data-stu-id="10b03-102">Loading the DOM and runtime environment</span></span>



<span data-ttu-id="10b03-103">外接程序在运行自己的自定义逻辑前必须确保 DOM 和 Office 外接程序运行时环境都已加载。</span><span class="sxs-lookup"><span data-stu-id="10b03-103">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span> 

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="10b03-104">启动内容或任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="10b03-104">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="10b03-105">下图显示了在 Excel、PowerPoint、Project、Word 或 Access 中启动内容或任务窗格外接程序所涉及的事件流。</span><span class="sxs-lookup"><span data-stu-id="10b03-105">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, Word, or Access.</span></span>

![启动内容/任务窗格外接程序时的事件流](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="10b03-107">启动内容/任务窗格外接程序时，将发生以下事件：</span><span class="sxs-lookup"><span data-stu-id="10b03-107">The following events occur when a content or task pane add-in starts:</span></span> 



1. <span data-ttu-id="10b03-108">用户打开已包含加载项的文档，或在文档中插入加载项。</span><span class="sxs-lookup"><span data-stu-id="10b03-108">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>
    
2. <span data-ttu-id="10b03-109">Office 主机应用从 AppSource、SharePoint 上的加载项目录或源自的共享文件夹目录中读取加载项 XML 清单。</span><span class="sxs-lookup"><span data-stu-id="10b03-109">The Office host application reads the add-in's XML manifest from AppSource, an add-in catalog on SharePoint, or the shared folder catalog it originates from.</span></span>
    
3. <span data-ttu-id="10b03-110">Office 主机应用在浏览器控件中打开加载项的 HTML 页面。</span><span class="sxs-lookup"><span data-stu-id="10b03-110">The Office host application opens the add-in's HTML page in a browser control.</span></span>
    
    <span data-ttu-id="10b03-p101">后面的两个步骤第 4 步和第 5 步以异步方式并行发生。因此，您的加载项代码必须在继续之前确保 DOM 和加载项运行时环境已加载完。</span><span class="sxs-lookup"><span data-stu-id="10b03-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>
    
4. <span data-ttu-id="10b03-113">浏览器控件加载 DOM 和 HTML 正文，并调用 **window.onload** 事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="10b03-113">The browser control loads the DOM and HTML body, and calls the event handler for the  **window.onload** event.</span></span>
    
5. <span data-ttu-id="10b03-114">Office 主机应用程序加载运行时环境，这将从内容分发网络 (CDN) 服务器中为 JavaScript 库文件下载并缓存 JavaScript API，然后为 [Office](/javascript/api/office) 对象的 [initialize](/javascript/api/office#initialize-reason-) 事件调用加载项的事件处理程序（如果已为其分配处理程序）。</span><span class="sxs-lookup"><span data-stu-id="10b03-114">The Office host application loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](/javascript/api/office#initialize-reason-) event of the [Office](/javascript/api/office) object, if a handler has been assigned to it.</span></span> <span data-ttu-id="10b03-115">此时它还会检查是否有任何回调（或链接 `then()` 函数）已传递（或链接）到 `Office.onReady` 处理程序。</span><span class="sxs-lookup"><span data-stu-id="10b03-115">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="10b03-116">有关 `Office.initialize` 与 `Office.onReady` 之间的区别的详细信息，请参阅[初始化加载项](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in)。</span><span class="sxs-lookup"><span data-stu-id="10b03-116">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initializing your add-in](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in).</span></span>
    
6. <span data-ttu-id="10b03-117">当 DOM 和 HTML 正文加载完毕并且加载项完成初始化后，加载项的主函数就可以继续进行。</span><span class="sxs-lookup"><span data-stu-id="10b03-117">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>
    

## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="10b03-118">启动 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="10b03-118">Startup of an Outlook add-in</span></span>



<span data-ttu-id="10b03-119">下图显示了启动在台式机、平板电脑或智能手机上运行的 Outlook 外接程序所涉及的事件流。</span><span class="sxs-lookup"><span data-stu-id="10b03-119">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![启动 Outlook 外接程序时的事件流](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="10b03-121">启动 Outlook 外接程序时，将发生以下事件：</span><span class="sxs-lookup"><span data-stu-id="10b03-121">The following events occur when an Outlook add-in starts:</span></span> 



1. <span data-ttu-id="10b03-122">当 Outlook 启动时，Outlook 读取已为用户的电子邮件帐户安装的 Outlook 外接程序的 XML 清单。</span><span class="sxs-lookup"><span data-stu-id="10b03-122">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>
    
2. <span data-ttu-id="10b03-123">用户选择 Outlook 中的一个项目。</span><span class="sxs-lookup"><span data-stu-id="10b03-123">The user selects an item in Outlook.</span></span>
    
3. <span data-ttu-id="10b03-124">如果所选项目满足某个 Outlook 外接程序的激活条件，则 Outlook 将激活该外接程序，并使其按钮在 UI 中可见。</span><span class="sxs-lookup"><span data-stu-id="10b03-124">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>
    
4. <span data-ttu-id="10b03-p103">如果用户单击该按钮以启动 Outlook 外接程序，Outlook 将在浏览器控件中打开 HTML 页面。下面两个步骤（步骤 5 和 6）并行发生。</span><span class="sxs-lookup"><span data-stu-id="10b03-p103">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>
    
5. <span data-ttu-id="10b03-127">浏览器控件加载 DOM 和 HTML 正文，并调用 **onload** 事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="10b03-127">The browser control loads the DOM and HTML body, and calls the event handler for the  **onload** event.</span></span>
    
6. <span data-ttu-id="10b03-128">Outlook 加载运行时环境，这将从内容分发网络 (CDN) 服务器中为 JavaScript 库文件下载并缓存 JavaScript API，然后为 [Office](/javascript/api/office) 加载项对象的 [initialize](/javascript/api/office#initialize-reason-) 事件调用事件处理程序（如果已为其分配处理程序）。</span><span class="sxs-lookup"><span data-stu-id="10b03-128">Outlook loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the event handler for the [initialize](/javascript/api/office#initialize-reason-) event of the [Office](/javascript/api/office) object of the add-in, if a handler has been assigned to it.</span></span> <span data-ttu-id="10b03-129">此时它还会检查是否有任何回调（或链接 `then()` 函数）已传递（或链接）到 `Office.onReady` 处理程序。</span><span class="sxs-lookup"><span data-stu-id="10b03-129">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="10b03-130">有关 `Office.initialize` 与 `Office.onReady` 之间的区别的详细信息，请参阅[初始化加载项](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in)。</span><span class="sxs-lookup"><span data-stu-id="10b03-130">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initializing your add-in](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in).</span></span>
    
7. <span data-ttu-id="10b03-131">当 DOM 和 HTML 正文加载完毕并且加载项完成初始化后，加载项的主函数就可以继续进行。</span><span class="sxs-lookup"><span data-stu-id="10b03-131">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>
    

## <a name="checking-the-load-status"></a><span data-ttu-id="10b03-132">检查加载状态</span><span class="sxs-lookup"><span data-stu-id="10b03-132">Checking the load status</span></span>

<span data-ttu-id="10b03-133">检查 DOM 和运行时环境是否已完成加载的一种方法是使用 jQuery [.ready()](https://api.jquery.com/ready/) 函数：`$(document).ready()`。</span><span class="sxs-lookup"><span data-stu-id="10b03-133">One way to check that both the DOM and the runtime environment have finished loading is to use the jQuery [.ready()](https://api.jquery.com/ready/) function: `$(document).ready()`.</span></span> <span data-ttu-id="10b03-134">例如，以下 **onReady** 事件处理程序确保在专用于初始化加载项的代码运行之前先加载 DOM。</span><span class="sxs-lookup"><span data-stu-id="10b03-134">For example, the following **onReady** event handler makes sure the DOM is first loaded before the code specific to initializing the add-in runs.</span></span> <span data-ttu-id="10b03-135">随后，**onReady** 处理程序将继续使用 [mailbox.item](https://docs.microsoft.com/javascript/api/outlook/office.mailbox) 属性来获取当前在 Outlook 中选择的项目，并调用加载项的主函数，即 `initDialer`。</span><span class="sxs-lookup"><span data-stu-id="10b03-135">Subsequently, the **onReady** handler proceeds to use the [mailbox.item](https://docs.microsoft.com/javascript/api/outlook/office.mailbox) property to obtain the currently selected item in Outlook, and calls the main function of the add-in, `initDialer`.</span></span>

```js
Office.onReady()
    .then(
        // Checks for the DOM to load.
        $(document).ready(function () {
            // After the DOM is loaded, add-in-specific code can run.
            var mailbox = Office.context.mailbox;
            _Item = mailbox.item;
            initDialer();
        });
);
```

<span data-ttu-id="10b03-136">或者，你也可以在 **initialize** 事件处理程序中使用相同的代码，如下面的示例所示。</span><span class="sxs-lookup"><span data-stu-id="10b03-136">Alternatively, you can use the same code in an  **initialize** event handler as shown in the following example.</span></span>

```js
Office.initialize = function () {
    // Checks for the DOM to load.
    $(document).ready(function () {
        // After the DOM is loaded, add-in-specific code can run.
        var mailbox = Office.context.mailbox;
        _Item = mailbox.item;
        initDialer();
    });
}
```

<span data-ttu-id="10b03-137">这种方法可在任何 Office 加载项的 **onReady** 或 **initialize** 处理程序中使用。</span><span class="sxs-lookup"><span data-stu-id="10b03-137">This same technique can be used in the **onReady** or **initialize** handlers of any Office Add-in.</span></span>

<span data-ttu-id="10b03-138">电话拨号器示例 Outlook 加载项展示了略为不同的方法，此方法仅使用 JavaScript 检查这些相同条件。</span><span class="sxs-lookup"><span data-stu-id="10b03-138">The phone dialer sample Outlook add-in shows a slightly different approach using only JavaScript to check these same conditions.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="10b03-139">即使加载项没有初始化任务要执行，也必须至少包含对 **Office.onReady** 的调用或分配最简单的 **Office.initialize** 事件处理程序函数，如下面的示例所示。</span><span class="sxs-lookup"><span data-stu-id="10b03-139">Even if your add-in has no initialization tasks to perform, you must include at least a call of **Office.onReady** or assign minimal **Office.initialize** event handler function as shown in the following examples.</span></span>
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```
>
> <span data-ttu-id="10b03-140">如果未调用 **Office.onReady** 或分配 **Office.initialize** 事件处理程序，则加载项在启动时可能会引发错误。</span><span class="sxs-lookup"><span data-stu-id="10b03-140">If you do not call **Office.onReady** or assign an  **Office.initialize** event handler, your add-in may raise an error when it starts.</span></span> <span data-ttu-id="10b03-141">而且，如果某个用户尝试通过 Office Online Web 客户端（例如 Excel Online、PowerPoint Online 或 Outlook Web App）使用你的外接程序，则外接程序会无法运行。</span><span class="sxs-lookup"><span data-stu-id="10b03-141">Also, if a user attempts to use your add-in with an Office Online web client, such as Excel Online, PowerPoint Online, or Outlook Web App, it will fail to run.</span></span>
>
> <span data-ttu-id="10b03-142">如果你的加载项包括多个页，则在每次加载新页时，该页面必须调用 **Office.onReady** 或分配 **Office.initialize** 事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="10b03-142">If your add-in includes more than one page, whenever it loads a new page that page must either call **Office.onReady** or assign an  **Office.initialize** event handler.</span></span>

## <a name="see-also"></a><span data-ttu-id="10b03-143">另请参阅</span><span class="sxs-lookup"><span data-stu-id="10b03-143">See also</span></span>

- [<span data-ttu-id="10b03-144">了解适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="10b03-144">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
    

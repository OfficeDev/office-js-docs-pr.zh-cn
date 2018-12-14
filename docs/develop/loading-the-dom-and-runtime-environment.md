---
title: 加载 DOM 和运行时环境
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: ac4d26d964f844f08e1d2975c1be8bbccf40349f
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/14/2018
ms.locfileid: "27271060"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="2e811-102">加载 DOM 和运行时环境</span><span class="sxs-lookup"><span data-stu-id="2e811-102">Loading the DOM and runtime environment</span></span>



<span data-ttu-id="2e811-103">外接程序在运行自己的自定义逻辑前必须确保 DOM 和 Office 外接程序运行时环境都已加载。</span><span class="sxs-lookup"><span data-stu-id="2e811-103">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span> 

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="2e811-104">启动内容或任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="2e811-104">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="2e811-105">下图显示了在 Excel、PowerPoint、Project、Word 或 Access 中启动内容或任务窗格外接程序所涉及的事件流。</span><span class="sxs-lookup"><span data-stu-id="2e811-105">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, Word, or Access.</span></span>

![启动内容/任务窗格外接程序时的事件流](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="2e811-107">启动内容/任务窗格外接程序时，将发生以下事件：</span><span class="sxs-lookup"><span data-stu-id="2e811-107">The following events occur when a content or task pane add-in starts:</span></span> 



1. <span data-ttu-id="2e811-108">用户打开已包含加载项的文档，或在文档中插入加载项。</span><span class="sxs-lookup"><span data-stu-id="2e811-108">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>
    
2. <span data-ttu-id="2e811-109">Office 主机应用从 AppSource、SharePoint 上的加载项目录或源自的共享文件夹目录中读取加载项 XML 清单。</span><span class="sxs-lookup"><span data-stu-id="2e811-109">The Office host application reads the add-in's XML manifest from AppSource, an add-in catalog on SharePoint, or the shared folder catalog it originates from.</span></span>
    
3. <span data-ttu-id="2e811-110">Office 主机应用在浏览器控件中打开加载项的 HTML 页面。</span><span class="sxs-lookup"><span data-stu-id="2e811-110">The Office host application opens the add-in's HTML page in a browser control.</span></span>
    
    <span data-ttu-id="2e811-p101">后面的两个步骤第 4 步和第 5 步以异步方式并行发生。因此，您的加载项代码必须在继续之前确保 DOM 和加载项运行时环境已加载完。</span><span class="sxs-lookup"><span data-stu-id="2e811-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>
    
4. <span data-ttu-id="2e811-113">浏览器控件加载 DOM 和 HTML 正文，并调用  **window.onload** 事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="2e811-113">The browser control loads the DOM and HTML body, and calls the event handler for the  **window.onload** event.</span></span>
    
5. <span data-ttu-id="2e811-114">Office 主机应用程序加载运行时环境，这将从内容分发网络 (CDN) 服务器中为 JavaScript 库文件下载并缓存 JavaScript API，然后为 [Office](https://docs.microsoft.com/javascript/api/office?view=office-js) 对象的 [initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) 事件调用加载项的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="2e811-114">The Office host application loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) event of the [Office](https://docs.microsoft.com/javascript/api/office?view=office-js) object.</span></span>
    
6. <span data-ttu-id="2e811-115">当 DOM 和 HTML 正文加载完毕并且加载项完成初始化后，加载项的主函数就可以继续进行。</span><span class="sxs-lookup"><span data-stu-id="2e811-115">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>
    

## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="2e811-116">启动 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="2e811-116">Startup of an Outlook add-in</span></span>



<span data-ttu-id="2e811-117">下图显示了启动在台式机、平板电脑或智能手机上运行的 Outlook 外接程序所涉及的事件流。</span><span class="sxs-lookup"><span data-stu-id="2e811-117">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![启动 Outlook 外接程序时的事件流](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="2e811-119">启动 Outlook 外接程序时，将发生以下事件：</span><span class="sxs-lookup"><span data-stu-id="2e811-119">The following events occur when an Outlook add-in starts:</span></span> 



1. <span data-ttu-id="2e811-120">当 Outlook 启动时，Outlook 读取已为用户的电子邮件帐户安装的 Outlook 外接程序的 XML 清单。</span><span class="sxs-lookup"><span data-stu-id="2e811-120">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>
    
2. <span data-ttu-id="2e811-121">用户选择 Outlook 中的一个项目。</span><span class="sxs-lookup"><span data-stu-id="2e811-121">The user selects an item in Outlook.</span></span>
    
3. <span data-ttu-id="2e811-122">如果所选项目满足某个 Outlook 外接程序的激活条件，则 Outlook 将激活该外接程序，并使其按钮在 UI 中可见。</span><span class="sxs-lookup"><span data-stu-id="2e811-122">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>
    
4. <span data-ttu-id="2e811-p102">如果用户单击该按钮以启动 Outlook 外接程序，Outlook 将在浏览器控件中打开 HTML 页面。下面两个步骤（步骤 5 和 6）并行发生。</span><span class="sxs-lookup"><span data-stu-id="2e811-p102">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>
    
5. <span data-ttu-id="2e811-125">浏览器控件加载 DOM 和 HTML 正文，并调用  **onload** 事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="2e811-125">The browser control loads the DOM and HTML body, and calls the event handler for the  **onload** event.</span></span>
    
6. <span data-ttu-id="2e811-126">Outlook 调用加载项的 [Office](https://docs.microsoft.com/javascript/api/office?view=office-js) 对象的 [initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) 事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="2e811-126">Outlook calls the event handler for the [initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) event of the [Office](https://docs.microsoft.com/javascript/api/office?view=office-js) object of the add-in.</span></span>
    
7. <span data-ttu-id="2e811-127">当 DOM 和 HTML 正文加载完毕并且加载项完成初始化后，加载项的主函数就可以继续进行。</span><span class="sxs-lookup"><span data-stu-id="2e811-127">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>
    

## <a name="checking-the-load-status"></a><span data-ttu-id="2e811-128">检查加载状态</span><span class="sxs-lookup"><span data-stu-id="2e811-128">Checking the load status</span></span>


<span data-ttu-id="2e811-p103">检查 DOM 和 运行时环境是否加载完毕的一种方式是使用 jQuery [.ready()](https://api.jquery.com/ready/) 函数： `$(document).ready()`。例如，以下  **initialize** 事件处理程序函数可确保在专门用于初始化外接程序的代码运行前先加载 DOM。随后， **initialize** 事件处理程序继续使用 [mailbox.item](https://docs.microsoft.com/javascript/api/outlook/office.mailbox?view=office-js) 属性获取 Outlook 中当前选定的项目，并调用外接程序的主函数 `initDialer`。</span><span class="sxs-lookup"><span data-stu-id="2e811-p103">One way to check that both the DOM and the runtime environment have finished loading is to use the jQuery [.ready()](https://api.jquery.com/ready/) function: `$(document).ready()`. For example, the following  **initialize** event handler function makes sure the DOM is first loaded before the code specific to initializing the add-in runs. Subsequently, the **initialize** event handler proceeds to use the [mailbox.item](https://docs.microsoft.com/javascript/api/outlook/office.mailbox?view=office-js) property to obtain the currently selected item in Outlook, and calls the main function of the add-in, `initDialer`.</span></span>


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

<span data-ttu-id="2e811-132">这种方法可在任何 Office 外接程序的  **initialize** 处理程序中使用。</span><span class="sxs-lookup"><span data-stu-id="2e811-132">This same technique can be used in the  **initialize** handler of any Office Add-in.</span></span>

<span data-ttu-id="2e811-133">电话拨号器示例 Outlook 加载项展示了略为不同的方法，此方法仅使用 JavaScript 检查这些相同条件。</span><span class="sxs-lookup"><span data-stu-id="2e811-133">The phone dialer sample Outlook add-in shows a slightly different approach using only JavaScript to check these same conditions.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="2e811-134">即使加载项没有初始化任务要执行，也必须至少添加最简单的 **Office.initialize** 事件处理程序函数，如下面的示例所示。</span><span class="sxs-lookup"><span data-stu-id="2e811-134">Even if your add-in has no initialization tasks to perform, you must include at least a minimal **Office.initialize** event handler function like the following example.</span></span>

```js
Office.initialize = function () {
};
```

<span data-ttu-id="2e811-p104">如果您无法加入  **Office.initialize** 事件处理程序，则启动加载项时可能会出错。此外，如果用户尝试将您的加载项与 Office Online Web 客户端（如 Excel Online、PowerPoint Online 或 Outlook Web App）结合使用，应用程序将无法运行。</span><span class="sxs-lookup"><span data-stu-id="2e811-p104">If you fail to include an  **Office.initialize** event handler, your add-in may raise an error when it starts. Also, if a user attempts to use your add-in with an Office Online web client, such as Excel Online, PowerPoint Online, or Outlook Web App, it will fail to run.</span></span>

<span data-ttu-id="2e811-137">如果您的加载项包括多个页，则在每次加载新页时，页面必须加入或调用  **Office.initialize** 事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="2e811-137">If your add-in includes more than one page, whenever it loads a new page that page must include or call an  **Office.initialize** event handler.</span></span>


## <a name="see-also"></a><span data-ttu-id="2e811-138">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2e811-138">See also</span></span>

- [<span data-ttu-id="2e811-139">了解适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="2e811-139">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
    

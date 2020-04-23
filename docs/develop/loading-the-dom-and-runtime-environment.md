---
title: 加载 DOM 和运行时环境
description: 加载 DOM 和 Office 外接程序运行时环境
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: 7248f5b09a54552c3f16a9bc97bd4eae9795c8cd
ms.sourcegitcommit: 9da68c00ecc00a2f307757e0f5a903a8e31b7769
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/22/2020
ms.locfileid: "43785715"
---
# <a name="loading-the-dom-and-runtime-environment"></a><span data-ttu-id="a5425-103">加载 DOM 和运行时环境</span><span class="sxs-lookup"><span data-stu-id="a5425-103">Loading the DOM and runtime environment</span></span>

<span data-ttu-id="a5425-104">外接程序在运行自己的自定义逻辑前必须确保 DOM 和 Office 外接程序运行时环境都已加载。</span><span class="sxs-lookup"><span data-stu-id="a5425-104">An add-in must ensure that both the DOM and the Office Add-ins runtime environment are loaded before running its own custom logic.</span></span>

## <a name="startup-of-a-content-or-task-pane-add-in"></a><span data-ttu-id="a5425-105">启动内容或任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="a5425-105">Startup of a content or task pane add-in</span></span>

<span data-ttu-id="a5425-106">下图显示了在 Excel、PowerPoint、Project 或 Word 中启动内容或任务窗格加载项所涉及的事件流。</span><span class="sxs-lookup"><span data-stu-id="a5425-106">The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, or Word.</span></span>

![启动内容/任务窗格外接程序时的事件流](../images/office15-app-sdk-loading-dom-agave-runtime.png)

<span data-ttu-id="a5425-108">启动内容/任务窗格外接程序时，将发生以下事件：</span><span class="sxs-lookup"><span data-stu-id="a5425-108">The following events occur when a content or task pane add-in starts:</span></span>

1. <span data-ttu-id="a5425-109">用户打开已包含加载项的文档，或在文档中插入加载项。</span><span class="sxs-lookup"><span data-stu-id="a5425-109">The user opens a document that already contains an add-in or inserts an add-in in the document.</span></span>

2. <span data-ttu-id="a5425-110">Office 主机应用从 AppSource、SharePoint 上的应用目录或源自的共享文件夹目录中读取加载项 XML 清单。</span><span class="sxs-lookup"><span data-stu-id="a5425-110">The Office host application reads the add-in's XML manifest from AppSource, an app catalog on SharePoint, or the shared folder catalog it originates from.</span></span>

3. <span data-ttu-id="a5425-111">Office 主机应用在浏览器控件中打开加载项的 HTML 页面。</span><span class="sxs-lookup"><span data-stu-id="a5425-111">The Office host application opens the add-in's HTML page in a browser control.</span></span>

    <span data-ttu-id="a5425-p101">后面的两个步骤第 4 步和第 5 步以异步方式并行发生。因此，您的加载项代码必须在继续之前确保 DOM 和加载项运行时环境已加载完。</span><span class="sxs-lookup"><span data-stu-id="a5425-p101">The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.</span></span>

4. <span data-ttu-id="a5425-114">浏览器控件加载 DOM 和 HTML 正文，并调用`window.onload`事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="a5425-114">The browser control loads the DOM and HTML body, and calls the event handler for the `window.onload` event.</span></span>

5. <span data-ttu-id="a5425-115">Office 主机应用程序加载运行时环境，这将从内容分发网络 (CDN) 服务器中为 JavaScript 库文件下载并缓存 JavaScript API，然后为 [Office](/javascript/api/office) 对象的 [initialize](/javascript/api/office#office-initialize-reason-) 事件调用加载项的事件处理程序（如果已为其分配处理程序）。</span><span class="sxs-lookup"><span data-stu-id="a5425-115">The Office host application loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object, if a handler has been assigned to it.</span></span> <span data-ttu-id="a5425-116">此时它还会检查是否有任何回调（或链接 `then()` 函数）已传递（或链接）到 `Office.onReady` 处理程序。</span><span class="sxs-lookup"><span data-stu-id="a5425-116">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="a5425-117">有关`Office.initialize`和`Office.onReady`的区别的详细信息，请参阅[初始化外接程序](initialize-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="a5425-117">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

6. <span data-ttu-id="a5425-118">当 DOM 和 HTML 正文加载完毕并且加载项完成初始化后，加载项的主函数就可以继续进行。</span><span class="sxs-lookup"><span data-stu-id="a5425-118">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="startup-of-an-outlook-add-in"></a><span data-ttu-id="a5425-119">启动 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="a5425-119">Startup of an Outlook add-in</span></span>

<span data-ttu-id="a5425-120">下图显示了启动在台式机、平板电脑或智能手机上运行的 Outlook 外接程序所涉及的事件流。</span><span class="sxs-lookup"><span data-stu-id="a5425-120">The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.</span></span>

![启动 Outlook 外接程序时的事件流](../images/outlook15-loading-dom-agave-runtime.png)

<span data-ttu-id="a5425-122">启动 Outlook 外接程序时，将发生以下事件：</span><span class="sxs-lookup"><span data-stu-id="a5425-122">The following events occur when an Outlook add-in starts:</span></span>

1. <span data-ttu-id="a5425-123">当 Outlook 启动时，Outlook 读取已为用户的电子邮件帐户安装的 Outlook 外接程序的 XML 清单。</span><span class="sxs-lookup"><span data-stu-id="a5425-123">When Outlook starts, Outlook reads the XML manifests for Outlook add-ins that have been installed for the user's email account.</span></span>

2. <span data-ttu-id="a5425-124">用户选择 Outlook 中的一个项目。</span><span class="sxs-lookup"><span data-stu-id="a5425-124">The user selects an item in Outlook.</span></span>

3. <span data-ttu-id="a5425-125">如果所选项目满足某个 Outlook 外接程序的激活条件，则 Outlook 将激活该外接程序，并使其按钮在 UI 中可见。</span><span class="sxs-lookup"><span data-stu-id="a5425-125">If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.</span></span>

4. <span data-ttu-id="a5425-p103">如果用户单击该按钮以启动 Outlook 外接程序，Outlook 将在浏览器控件中打开 HTML 页面。下面两个步骤（步骤 5 和 6）并行发生。</span><span class="sxs-lookup"><span data-stu-id="a5425-p103">If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a browser control. The next two steps, steps 5 and 6, occur in parallel.</span></span>

5. <span data-ttu-id="a5425-128">浏览器控件加载 DOM 和 HTML 正文，并调用`onload`事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="a5425-128">The browser control loads the DOM and HTML body, and calls the event handler for the `onload` event.</span></span>

6. <span data-ttu-id="a5425-129">Outlook 加载运行时环境，这将从内容分发网络 (CDN) 服务器中为 JavaScript 库文件下载并缓存 JavaScript API，然后为 [Office](/javascript/api/office) 加载项对象的 [initialize](/javascript/api/office#office-initialize-reason-) 事件调用事件处理程序（如果已为其分配处理程序）。</span><span class="sxs-lookup"><span data-stu-id="a5425-129">Outlook loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the event handler for the [initialize](/javascript/api/office#office-initialize-reason-) event of the [Office](/javascript/api/office) object of the add-in, if a handler has been assigned to it.</span></span> <span data-ttu-id="a5425-130">此时它还会检查是否有任何回调（或链接 `then()` 函数）已传递（或链接）到 `Office.onReady` 处理程序。</span><span class="sxs-lookup"><span data-stu-id="a5425-130">At this time it also checks to see if any callbacks (or chained `then()` functions) have been passed (or chained) to the `Office.onReady` handler.</span></span> <span data-ttu-id="a5425-131">有关`Office.initialize`和`Office.onReady`的区别的详细信息，请参阅[初始化外接程序](initialize-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="a5425-131">For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).</span></span>

7. <span data-ttu-id="a5425-132">当 DOM 和 HTML 正文加载完毕并且加载项完成初始化后，加载项的主函数就可以继续进行。</span><span class="sxs-lookup"><span data-stu-id="a5425-132">When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.</span></span>


## <a name="checking-the-load-status"></a><span data-ttu-id="a5425-133">检查加载状态</span><span class="sxs-lookup"><span data-stu-id="a5425-133">Checking the load status</span></span>

<span data-ttu-id="a5425-134">检查 DOM 和运行时环境是否已完成加载的一种方法是使用 jQuery [.ready()](https://api.jquery.com/ready/) 函数：`$(document).ready()`。</span><span class="sxs-lookup"><span data-stu-id="a5425-134">One way to check that both the DOM and the runtime environment have finished loading is to use the jQuery [.ready()](https://api.jquery.com/ready/) function: `$(document).ready()`.</span></span> <span data-ttu-id="a5425-135">例如，以下`onReady`事件处理程序确保在初始化外接程序的特定代码运行之前先加载 DOM。</span><span class="sxs-lookup"><span data-stu-id="a5425-135">For example, the following `onReady` event handler makes sure the DOM is first loaded before the code specific to initializing the add-in runs.</span></span> <span data-ttu-id="a5425-136">随后， `onReady`处理程序将继续使用[邮箱. Item](/javascript/api/outlook/office.mailbox#item)属性获取 Outlook 中当前选定的项，并调用外接程序的主函数。 `initDialer`</span><span class="sxs-lookup"><span data-stu-id="a5425-136">Subsequently, the `onReady` handler proceeds to use the [mailbox.item](/javascript/api/outlook/office.mailbox#item) property to obtain the currently selected item in Outlook, and calls the main function of the add-in, `initDialer`.</span></span>

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

<span data-ttu-id="a5425-137">或者，您可以在`initialize`事件处理程序中使用相同的代码，如下面的示例所示。</span><span class="sxs-lookup"><span data-stu-id="a5425-137">Alternatively, you can use the same code in an `initialize` event handler as shown in the following example.</span></span>

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

<span data-ttu-id="a5425-138">此方法可在任何 Office 外接`onReady`程序`initialize`的或处理程序中使用。</span><span class="sxs-lookup"><span data-stu-id="a5425-138">This same technique can be used in the `onReady` or `initialize` handlers of any Office Add-in.</span></span>

<span data-ttu-id="a5425-139">电话拨号器示例 Outlook 加载项展示了略为不同的方法，此方法仅使用 JavaScript 检查这些相同条件。</span><span class="sxs-lookup"><span data-stu-id="a5425-139">The phone dialer sample Outlook add-in shows a slightly different approach using only JavaScript to check these same conditions.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a5425-140">即使外接程序没有要执行的初始化任务，也必须至少包含调用`Office.onReady`或分配最少`Office.initialize`事件处理程序函数，如以下示例中所示。</span><span class="sxs-lookup"><span data-stu-id="a5425-140">Even if your add-in has no initialization tasks to perform, you must include at least a call of `Office.onReady` or assign minimal `Office.initialize` event handler function as shown in the following examples.</span></span>
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```
>
> <span data-ttu-id="a5425-141">如果您不调用`Office.onReady`或分配`Office.initialize`事件处理程序，则加载项会在启动时引发错误。</span><span class="sxs-lookup"><span data-stu-id="a5425-141">If you do not call `Office.onReady` or assign an `Office.initialize` event handler, your add-in may raise an error when it starts.</span></span> <span data-ttu-id="a5425-142">而且，如果某个用户尝试通过 Office Web 客户端（例如 Excel、PowerPoint 或 Outlook）使用你的加载项，则加载项会无法运行。</span><span class="sxs-lookup"><span data-stu-id="a5425-142">Also, if a user attempts to use your add-in with an Office web client, such as Excel, PowerPoint, or Outlook, it will fail to run.</span></span>
>
> <span data-ttu-id="a5425-143">如果您的外接程序包含多个页面，则每当它加载一个新页面时，该页都`Office.onReady`必须调用或`Office.initialize`分配一个事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="a5425-143">If your add-in includes more than one page, whenever it loads a new page that page must either call `Office.onReady` or assign an `Office.initialize` event handler.</span></span>

## <a name="see-also"></a><span data-ttu-id="a5425-144">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a5425-144">See also</span></span>

- [<span data-ttu-id="a5425-145">了解 Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="a5425-145">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="a5425-146">初始化 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="a5425-146">Initialize your Office Add-in</span></span>](initialize-add-in.md)

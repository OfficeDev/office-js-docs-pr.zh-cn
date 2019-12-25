---
title: OneNote JavaScript API 编程概述
description: ''
ms.date: 07/05/2019
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 9724de8c25a535884c4700a165e661028aee6608
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851507"
---
# <a name="onenote-javascript-api-programming-overview"></a><span data-ttu-id="9395b-102">OneNote JavaScript API 编程概述</span><span class="sxs-lookup"><span data-stu-id="9395b-102">OneNote JavaScript API programming overview</span></span>

<span data-ttu-id="9395b-103">OneNote 引入了适用于 OneNote 网页版加载项的 JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="9395b-103">OneNote introduces a JavaScript API for OneNote add-ins on the web.</span></span> <span data-ttu-id="9395b-104">可以创建任务窗格加载项、内容加载项，以及与 OneNote 对象交互并连接到 Web 服务或其他基于 Web 的资源的加载项命令。</span><span class="sxs-lookup"><span data-stu-id="9395b-104">You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.</span></span>

> [!NOTE]
> <span data-ttu-id="9395b-p102">如果计划将加载项[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](/office/dev/store/validation-policies)。例如，加载项必须适用于支持已定义方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性](../overview/office-add-in-availability.md)页面）。</span><span class="sxs-lookup"><span data-stu-id="9395b-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="9395b-107">Office 加载项的组件</span><span class="sxs-lookup"><span data-stu-id="9395b-107">Components of an Office Add-in</span></span>

<span data-ttu-id="9395b-108">加载项由两个基本部分组成：</span><span class="sxs-lookup"><span data-stu-id="9395b-108">Add-ins consist of two basic components:</span></span>

- <span data-ttu-id="9395b-109">包含网页和所有相应 JavaScript、CSS 或其他文件的 **Web 应用程序**。</span><span class="sxs-lookup"><span data-stu-id="9395b-109">A **web application** consisting of a webpage and any required JavaScript, CSS, or other files.</span></span> <span data-ttu-id="9395b-110">这些文件托管在 Web 服务器或 Web 托管服务上，例如 Microsoft Azure。</span><span class="sxs-lookup"><span data-stu-id="9395b-110">These files are hosted on a web server or web hosting service, such as Microsoft Azure.</span></span> <span data-ttu-id="9395b-111">在 OneNote 网页版中，Web 应用程序在浏览器控件或 iframe 中显示。</span><span class="sxs-lookup"><span data-stu-id="9395b-111">In OneNote on the web, the web application displays in a browser control or iframe.</span></span>

- <span data-ttu-id="9395b-p104">**XML 清单**指定外接程序网页的 URL 和适用于外接程序的任何访问要求、设置和功能。此文件存储在客户端上。OneNote 外接程序使用与其他 Office 外接程序相同的 [清单](../develop/add-in-manifests.md)格式。</span><span class="sxs-lookup"><span data-stu-id="9395b-p104">An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.</span></span>

<span data-ttu-id="9395b-115">**Office 加载项 = 清单 + 网页**</span><span class="sxs-lookup"><span data-stu-id="9395b-115">**Office Add-in = Manifest + Webpage**</span></span>

![Office 加载项包含清单和网页](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a><span data-ttu-id="9395b-117">使用 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="9395b-117">Using the JavaScript API</span></span>

<span data-ttu-id="9395b-p105">加载项使用托管应用程序的运行时上下文以访问 JavaScript API。API 有两层：</span><span class="sxs-lookup"><span data-stu-id="9395b-p105">Add-ins use the runtime context of the host application to access the JavaScript API. The API has two layers:</span></span> 

- <span data-ttu-id="9395b-120">用于执行 OneNote 专属操作的**主机特定 API**，可通过 **Application** 对象访问。</span><span class="sxs-lookup"><span data-stu-id="9395b-120">A **host-specific API** for OneNote-specific operations, accessed through the **Application** object.</span></span>
- <span data-ttu-id="9395b-121">跨 Office 应用程序分享的**通用 API**，通过**Document** 对象访问。</span><span class="sxs-lookup"><span data-stu-id="9395b-121">A **Common API** that's shared across Office applications, accessed through the **Document** object.</span></span>

### <a name="accessing-the-host-specific-api-through-the-application-object"></a><span data-ttu-id="9395b-122">通过 *Application* 对象访问主机特定 API。</span><span class="sxs-lookup"><span data-stu-id="9395b-122">Accessing the host-specific API through the *Application* object</span></span>

<span data-ttu-id="9395b-123">**Application** 对象可用于访问 OneNote 对象，如 **Notebook**、**Section** 和 **Page**。</span><span class="sxs-lookup"><span data-stu-id="9395b-123">Use the **Application** object to access OneNote objects such as **Notebook**, **Section**, and **Page**.</span></span> <span data-ttu-id="9395b-124">通过主机特定 API，可在代理对象上运行批处理操作。</span><span class="sxs-lookup"><span data-stu-id="9395b-124">With host-specific APIs, you run batch operations on proxy objects.</span></span> <span data-ttu-id="9395b-125">基本流程类似如下：</span><span class="sxs-lookup"><span data-stu-id="9395b-125">The basic flow goes something like this:</span></span> 

1. <span data-ttu-id="9395b-126">从上下文中获取应用程序实例。</span><span class="sxs-lookup"><span data-stu-id="9395b-126">Get the application instance from the context.</span></span>

2. <span data-ttu-id="9395b-p107">创建您想要使用的表示 OneNote 对象的代理。通过读取和写入代理对象的属性和调用其方法，您可以与其同步交互。</span><span class="sxs-lookup"><span data-stu-id="9395b-p107">Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.</span></span>

3. <span data-ttu-id="9395b-p108">调用代理上的 **load** 以使用在参数中指定的属性值填充它。此调用将添加至命令队列中。</span><span class="sxs-lookup"><span data-stu-id="9395b-p108">Call **load** on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.</span></span>

   > [!NOTE]
   > <span data-ttu-id="9395b-131">API 方法调用（如 `context.application.getActiveSection().pages;`）也会添加到队列中。</span><span class="sxs-lookup"><span data-stu-id="9395b-131">Method calls to the API (such as `context.application.getActiveSection().pages;`) are also added to the queue.</span></span>

4. <span data-ttu-id="9395b-p109">调用 **context.sync** 以按它们已排队的顺序运行所有排队的命令。这将同步您正在运行的脚本和真实对象之间的状态，并通过检索已加载的用于您的脚本的 OneNote 对象的属性实现。您可以使用返回的 promise 对象以链接其他操作。</span><span class="sxs-lookup"><span data-stu-id="9395b-p109">Call **context.sync** to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.</span></span>

<span data-ttu-id="9395b-135">例如：</span><span class="sxs-lookup"><span data-stu-id="9395b-135">For example:</span></span>

```js
function getPagesInSection() {
    OneNote.run(function (context) {

        // Get the pages in the current section.
        var pages = context.application.getActiveSection().pages;

        // Queue a command to load the id and title for each page.
        pages.load('id,title');

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync()
            .then(function () {

                // Read the id and title of each page.
                $.each(pages.items, function(index, page) {
                    var pageId = page.id;
                    var pageTitle = page.title;
                    console.log(pageTitle + ': ' + pageId);
                });
            })
            .catch(function (error) {
                app.showNotification("Error: " + error);
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    });
}
```

<span data-ttu-id="9395b-136">可以在 [API 参考](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference) 中找到受支持的 OneNote 对象和操作。</span><span class="sxs-lookup"><span data-stu-id="9395b-136">You can find supported OneNote objects and operations in the [API reference](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference).</span></span>

#### <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="9395b-137">OneNote JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="9395b-137">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="9395b-138">要求集是指各组已命名的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="9395b-138">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="9395b-139">Office 外接程序使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持外接程序所需的 API。</span><span class="sxs-lookup"><span data-stu-id="9395b-139">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="9395b-140">有关 OneNote JavaScript API 要求集的详细信息，请参阅 [OneNote JavaScript API 要求集](../reference/requirement-sets/onenote-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="9395b-140">For detailed information about OneNote JavaScript API requirement sets, see [OneNote JavaScript API requirement sets](../reference/requirement-sets/onenote-api-requirement-sets.md).</span></span>

### <a name="accessing-the-common-api-through-the-document-object"></a><span data-ttu-id="9395b-141">通过 *Document* 对象访问通用 API</span><span class="sxs-lookup"><span data-stu-id="9395b-141">Accessing the Common API through the *Document* object</span></span>

<span data-ttu-id="9395b-142">使用 **Document** 对象以访问通用 API，例如 [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) 和 [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 方法。</span><span class="sxs-lookup"><span data-stu-id="9395b-142">Use the **Document** object to access the Common API, such as the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) and [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) methods.</span></span> 


<span data-ttu-id="9395b-143">例如：</span><span class="sxs-lookup"><span data-stu-id="9395b-143">For example:</span></span>  

```js
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```

<span data-ttu-id="9395b-144">OneNote 加载项仅支持以下通用 API：</span><span class="sxs-lookup"><span data-stu-id="9395b-144">OneNote add-ins support only the following Common APIs:</span></span>

| <span data-ttu-id="9395b-145">API</span><span class="sxs-lookup"><span data-stu-id="9395b-145">API</span></span> | <span data-ttu-id="9395b-146">注释</span><span class="sxs-lookup"><span data-stu-id="9395b-146">Notes</span></span> |
|:------|:------|
| [<span data-ttu-id="9395b-147">Office.context.document.getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="9395b-147">Office.context.document.getSelectedDataAsync</span></span>](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) | <span data-ttu-id="9395b-148">仅限 **Office.CoercionType.Text** 和 **Office.CoercionType.Matrix**</span><span class="sxs-lookup"><span data-stu-id="9395b-148">**Office.CoercionType.Text** and **Office.CoercionType.Matrix** only</span></span> |
| [<span data-ttu-id="9395b-149">Office.context.document.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="9395b-149">Office.context.document.setSelectedDataAsync</span></span>](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) | <span data-ttu-id="9395b-150">仅限 **Office.CoercionType.Text**、**Office.CoercionType.Image** 和 **Office.CoercionType.Html**</span><span class="sxs-lookup"><span data-stu-id="9395b-150">**Office.CoercionType.Text**, **Office.CoercionType.Image**, and **Office.CoercionType.Html** only</span></span> | 
| [<span data-ttu-id="9395b-151">var mySetting = Office.context.document.settings.get(name);</span><span class="sxs-lookup"><span data-stu-id="9395b-151">var mySetting = Office.context.document.settings.get(name);</span></span>](/javascript/api/office/office.settings#get-name-) | <span data-ttu-id="9395b-152">设置仅受内容外接程序支持</span><span class="sxs-lookup"><span data-stu-id="9395b-152">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="9395b-153">Office.context.document.settings.set(name, value);</span><span class="sxs-lookup"><span data-stu-id="9395b-153">Office.context.document.settings.set(name, value);</span></span>](/javascript/api/office/office.settings#set-name--value-) | <span data-ttu-id="9395b-154">设置仅受内容外接程序支持</span><span class="sxs-lookup"><span data-stu-id="9395b-154">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="9395b-155">Office.EventType.DocumentSelectionChanged</span><span class="sxs-lookup"><span data-stu-id="9395b-155">Office.EventType.DocumentSelectionChanged</span></span>](/javascript/api/office/office.documentselectionchangedeventargs) ||

<span data-ttu-id="9395b-156">一般情况下，需要使用通用 API 执行主机特定 API 不支持的操作。</span><span class="sxs-lookup"><span data-stu-id="9395b-156">In general, you only use the Common API to do something that isn't supported in the host-specific API.</span></span> <span data-ttu-id="9395b-157">若要详细了解如何使用通用 API，请参阅 [Office JavaScript API 对象模型](../develop/office-javascript-api-object-model.md)。</span><span class="sxs-lookup"><span data-stu-id="9395b-157">To learn more about using the Common API, see [Office JavaScript API object model](../develop/office-javascript-api-object-model.md).</span></span>


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a><span data-ttu-id="9395b-158">OneNote 对象模型图</span><span class="sxs-lookup"><span data-stu-id="9395b-158">OneNote object model diagram</span></span> 
<span data-ttu-id="9395b-159">下图表示了 OneNote JavaScript API 中当前可用的内容。</span><span class="sxs-lookup"><span data-stu-id="9395b-159">The following diagram represents what's currently available in the OneNote JavaScript API.</span></span>

  ![OneNote 对象模型图](../images/onenote-om.png)


## <a name="see-also"></a><span data-ttu-id="9395b-161">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9395b-161">See also</span></span>

- [<span data-ttu-id="9395b-162">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="9395b-162">Building Office Add-ins using Office.js book</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="9395b-163">生成首个 OneNote 加载项</span><span class="sxs-lookup"><span data-stu-id="9395b-163">Build your first OneNote add-in</span></span>](../quickstarts/onenote-quickstart.md)
- [<span data-ttu-id="9395b-164">OneNote JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="9395b-164">OneNote JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="9395b-165">Rubric Grader 示例</span><span class="sxs-lookup"><span data-stu-id="9395b-165">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="9395b-166">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="9395b-166">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

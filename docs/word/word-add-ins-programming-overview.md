---
title: Word 加载项概述
description: 了解 Word 加载项的基本知识
ms.date: 07/28/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: b531ec5c2a5fa1e3e9366f703a57e815a5711b5a
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293070"
---
# <a name="word-add-ins-overview"></a><span data-ttu-id="d6205-103">Word 加载项概述</span><span class="sxs-lookup"><span data-stu-id="d6205-103">Word add-ins overview</span></span>

<span data-ttu-id="d6205-p101">要创建解决方案来扩展 Word 功能？例如，涉及自动文档程序集的解决方案？或从其他数据源绑定到并访问 Word 文档中数据的解决方案？可以使用 Office 加载项平台，其中包含 Word JavaScript API 和Office JavaScript API，可用于扩展在 Windows 桌面设备、Mac 或云中运行的 Word 客户端。</span><span class="sxs-lookup"><span data-stu-id="d6205-p101">Do you want to create a solution that extends the functionality of Word? For example, one that involves automated document assembly? Or a solution that binds to and accesses data in a Word document from other data sources? You can use the Office Add-ins platform, which includes the Word JavaScript API and the Office JavaScript API, to extend Word clients running on a Windows desktop, on a Mac, or in the cloud.</span></span>

<span data-ttu-id="d6205-p102">Word 外接程序是 [Office 外接程序平台](../overview/office-add-ins.md)上众多开发选项中的一项。外接程序命令可用于扩展 Word 用户界面并启动运行 JavaScript 并与 Word 文档中内容交互的任务窗格。在浏览器中可以运行的任何代码均可在 Word 外接程序中运行。与 Word 文档内容进行交互的外接程序可创建作用于 Word 对象的请求并同步对象状态。</span><span class="sxs-lookup"><span data-stu-id="d6205-p102">Word add-ins are one of the many development options that you have on the [Office Add-ins platform](../overview/office-add-ins.md). You can use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

<span data-ttu-id="d6205-112">下图中的示例展示了在任务窗格中运行的 Word 加载项。</span><span class="sxs-lookup"><span data-stu-id="d6205-112">The following figure shows an example of a Word add-in that runs in a task pane.</span></span>

<span data-ttu-id="d6205-113">*图 1：在 Word 的任务窗格中运行的加载项*</span><span class="sxs-lookup"><span data-stu-id="d6205-113">*Figure 1. Add-in running in a task pane in Word*</span></span>

![在 Word 的任务窗格中运行的外接程序](../images/word-add-in-show-host-client.png)

<span data-ttu-id="d6205-p103">Word 外接程序 (1) 可以将请求发送到 Word 文档 (2) 可以使用 JavaScript 来访问段落对象和更新、删除或移动段落。例如，下面的代码演示如何将一个新句子附加到该段落。</span><span class="sxs-lookup"><span data-stu-id="d6205-p103">The Word add-in (1) can send requests to the Word document (2) and can use JavaScript to access the paragraph object and update, delete, or move the paragraph. For example, the following code shows how to append a new sentence to that paragraph.</span></span>

```js
Word.run(function (context) {
    var paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load();
    return context.sync().then(function () {
        paragraphs.items[0].insertText(' New sentence in the paragraph.',
                                       Word.InsertLocation.end);
    }).then(context.sync);
});

```

<span data-ttu-id="d6205-p104">若要托管 Word 加载项，可以使用任何 Web 服务器技术（如 ASP.NET、NodeJS 或 Python）。可以使用常用的客户端框架（Ember、Backbone、Angular、React），也可以坚持使用 VanillaJS 开发解决方案，并能使用 Azure 等服务[验证](../develop/overview-authn-authz.md)和托管应用。</span><span class="sxs-lookup"><span data-stu-id="d6205-p104">You can use any web server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework -- Ember, Backbone, Angular, React -- or stick with VanillaJS to develop your solution, and you can use services like Azure to [authenticate](../develop/overview-authn-authz.md) and host your application.</span></span>

<span data-ttu-id="d6205-p105">通过 Word JavaScript API 可使应用程序访问 Word 文档中的对象和元数据。这些 API 可用于创建面向以下应用程序的外接程序：</span><span class="sxs-lookup"><span data-stu-id="d6205-p105">The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can use these APIs to create add-ins that target:</span></span>

* <span data-ttu-id="d6205-121">Windows 版 Word 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="d6205-121">Word 2013 or later on Windows</span></span>
* <span data-ttu-id="d6205-122">Word 网页版</span><span class="sxs-lookup"><span data-stu-id="d6205-122">Word on the web</span></span>
* <span data-ttu-id="d6205-123">Mac 版 Word 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="d6205-123">Word 2016 or later on Mac</span></span>
* <span data-ttu-id="d6205-124">iPad 版 Word</span><span class="sxs-lookup"><span data-stu-id="d6205-124">Word on iPad</span></span>

<span data-ttu-id="d6205-p106">加载项只需编写一次，即可跨多个平台在所有版本 Word 中运行。有关详细信息，请参阅 [Office 客户端应用程序和加载项平台可用性](../overview/office-add-in-availability.md)。</span><span class="sxs-lookup"><span data-stu-id="d6205-p106">Write your add-in once, and it will run in all versions of Word across multiple platforms. For details, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

## <a name="javascript-apis-for-word"></a><span data-ttu-id="d6205-127">适用于 Word 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="d6205-127">JavaScript APIs for Word</span></span>

<span data-ttu-id="d6205-128">有两组 JavaScript API 可用于与 Word 文档中的对象和元数据进行交互。</span><span class="sxs-lookup"><span data-stu-id="d6205-128">You can use two sets of JavaScript APIs to interact with the objects and metadata in a Word document.</span></span> <span data-ttu-id="d6205-129">第一组是在 Office 2013 中引入的[通用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="d6205-129">The first is the [Common API](/javascript/api/office), which was introduced in Office 2013.</span></span> <span data-ttu-id="d6205-130">通用 API 中的许多对象可以在由两个或多个 Office 客户端托管的加载项中使用。</span><span class="sxs-lookup"><span data-stu-id="d6205-130">Many of the objects in the Common API can be used in add-ins hosted by two or more Office clients.</span></span> <span data-ttu-id="d6205-131">此 API 广泛使用回调。</span><span class="sxs-lookup"><span data-stu-id="d6205-131">This API uses callbacks extensively.</span></span>

<span data-ttu-id="d6205-p108">第二组是 [Word JavaScript API](/javascript/api/word)。这是与 Word 2016 年一起引入的[应用程序特定 API 模型](../develop/application-specific-api-model.md)。它是强类型对象模型，可用于创建面向 Mac 版和 Windows 版 Word 2016 的 Word 加载项。此对象模型使用承诺模式，并提供对特定于 Word 的对象（如[正文](/javascript/api/word/word.body)、[内容控件](/javascript/api/word/word.contentcontrol)、[内联图片](/javascript/api/word/word.inlinepicture)和[段落](/javascript/api/word/word.paragraph)）的访问权限。Word JavaScript API 包括 TypeScript 定义和 vsdoc 文件，这样，你便可以在 IDE 中获得代码提示。</span><span class="sxs-lookup"><span data-stu-id="d6205-p108">The second is the [Word JavaScript API](/javascript/api/word). This is a [application-specific API model](../develop/application-specific-api-model.md) that was introduced with Word 2016. It's a strongly-typed object model that you can use to create Word add-ins that target Word 2016 on Mac and Windows. This object model uses promises and provides access to Word-specific objects like [body](/javascript/api/word/word.body), [content controls](/javascript/api/word/word.contentcontrol), [inline pictures](/javascript/api/word/word.inlinepicture), and [paragraphs](/javascript/api/word/word.paragraph). The Word JavaScript API includes TypeScript definitions and vsdoc files so that you can get code hints in your IDE.</span></span>

<span data-ttu-id="d6205-137">目前，所有 Word 客户端均支持共享的 Office JavaScript API，大多数客户端支持 Word JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="d6205-137">Currently, all Word clients support the shared Office JavaScript API, and most clients support the Word JavaScript API.</span></span> <span data-ttu-id="d6205-138">有关受支持的客户端的详细信息，请参阅 [Office 客户端应用程序和 Office 加载项的平台可用性](../overview/office-add-in-availability.md)。</span><span class="sxs-lookup"><span data-stu-id="d6205-138">For details about supported clients, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

<span data-ttu-id="d6205-p110">我们建议从 Word JavaScript API 开始，因为对象模型更易于使用。如果需要执行以下操作，请使用 Word JavaScript API：</span><span class="sxs-lookup"><span data-stu-id="d6205-p110">We recommend that you start with the Word JavaScript API because the object model is easier to use. Use the Word JavaScript API if you need to:</span></span>

* <span data-ttu-id="d6205-141">访问 Word 文档中的对象。</span><span class="sxs-lookup"><span data-stu-id="d6205-141">Access the objects in a Word document.</span></span>

<span data-ttu-id="d6205-142">在需要执行以下操作时，使用共享的 Office JavaScript API：</span><span class="sxs-lookup"><span data-stu-id="d6205-142">Use the shared Office JavaScript API when you need to:</span></span>

* <span data-ttu-id="d6205-143">面向 Word 2013。</span><span class="sxs-lookup"><span data-stu-id="d6205-143">Target Word 2013.</span></span>
* <span data-ttu-id="d6205-144">执行应用程序的初始操作。</span><span class="sxs-lookup"><span data-stu-id="d6205-144">Perform initial actions for the application.</span></span>
* <span data-ttu-id="d6205-145">检查支持的要求集。</span><span class="sxs-lookup"><span data-stu-id="d6205-145">Check the supported requirement set.</span></span>
* <span data-ttu-id="d6205-146">访问文档的元数据、设置和环境信息。</span><span class="sxs-lookup"><span data-stu-id="d6205-146">Access metadata, settings, and environmental information for the document.</span></span>
* <span data-ttu-id="d6205-147">绑定到文档中的部分并捕获事件。</span><span class="sxs-lookup"><span data-stu-id="d6205-147">Bind to sections in a document and capture events.</span></span>
* <span data-ttu-id="d6205-148">使用自定义 XML 部件。</span><span class="sxs-lookup"><span data-stu-id="d6205-148">Use custom XML parts.</span></span>
* <span data-ttu-id="d6205-149">打开一个对话框。</span><span class="sxs-lookup"><span data-stu-id="d6205-149">Open a dialog box.</span></span>

## <a name="next-steps"></a><span data-ttu-id="d6205-150">后续步骤</span><span class="sxs-lookup"><span data-stu-id="d6205-150">Next steps</span></span>

<span data-ttu-id="d6205-p111">准备好创建首个 Word 加载项了吗？请参阅[构建首个 Word 加载项](word-add-ins.md)。请使用[加载项清单](../develop/add-in-manifests.md)描述加载项的托管位置和显示方式，并定义权限和其他信息。</span><span class="sxs-lookup"><span data-stu-id="d6205-p111">Ready to create your first Word add-in? See [Build your first Word add-in](word-add-ins.md). Use the [add-in manifest](../develop/add-in-manifests.md) to describe where your add-in is hosted, how it is displayed, and define permissions and other information.</span></span>

<span data-ttu-id="d6205-154">若要了解如何设计世界一流的 Word 外接程序来为用户打造具有吸引力的体验，请参阅[设计指南](../design/add-in-design.md)和[最佳实践](../concepts/add-in-development-best-practices.md)。</span><span class="sxs-lookup"><span data-stu-id="d6205-154">To learn more about how to design a world class Word add-in that creates a compelling experience for your users, see [Design guidelines](../design/add-in-design.md) and [Best practices](../concepts/add-in-development-best-practices.md).</span></span>

<span data-ttu-id="d6205-155">开发加载项后，可以将它[发布](../publish/publish.md)到网络共享、应用目录或 AppSource。</span><span class="sxs-lookup"><span data-stu-id="d6205-155">After you develop your add-in, you can [publish](../publish/publish.md) it to a network share, an app catalog, or AppSource.</span></span>

## <a name="see-also"></a><span data-ttu-id="d6205-156">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d6205-156">See also</span></span>

* [<span data-ttu-id="d6205-157">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="d6205-157">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="d6205-158">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="d6205-158">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="d6205-159">Word JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="d6205-159">Word JavaScript API reference</span></span>](../reference/overview/word-add-ins-reference-overview.md)

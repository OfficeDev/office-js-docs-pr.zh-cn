---
title: 参考 Office JavaScript API 库
description: 了解如何在外接程序中引用 Office JavaScript API 库和类型定义。
ms.date: 06/23/2020
localization_priority: Normal
ms.openlocfilehash: 64dd08329b7bbc8c249bd270a431b6cbe93ec52c
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293182"
---
# <a name="referencing-the-office-javascript-api-library"></a><span data-ttu-id="fb264-103">参考 Office JavaScript API 库</span><span class="sxs-lookup"><span data-stu-id="fb264-103">Referencing the Office JavaScript API library</span></span>

<span data-ttu-id="fb264-104">[Office JAVASCRIPT API](../reference/javascript-api-for-office.md)库提供你的外接程序可用于与 Office 应用程序进行交互的 api。</span><span class="sxs-lookup"><span data-stu-id="fb264-104">The [Office JavaScript API](../reference/javascript-api-for-office.md) library provides the APIs that your add-in can use to interact with the Office application.</span></span> <span data-ttu-id="fb264-105">引用库的最简单方法是使用内容传递网络 (CDN) ，方法是 `<script>` 在 HTML 页面的部分中添加以下标记 `<head>` ：</span><span class="sxs-lookup"><span data-stu-id="fb264-105">The simplest way to reference the library is to use the content delivery network (CDN) by adding the following `<script>` tag within the `<head>` section of your HTML page:</span></span>  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

<span data-ttu-id="fb264-106">这将在您首次加载加载项时下载并缓存 Office JavaScript API 文件，以确保它使用的是最新的 Office.js 实现，并将其与指定版本相关联的文件。</span><span class="sxs-lookup"><span data-stu-id="fb264-106">This will download and cache the Office JavaScript API files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fb264-107">您必须从页面的部分中引用 Office JavaScript API， `<head>` 以确保 API 在任何 body 元素之前完全初始化。</span><span class="sxs-lookup"><span data-stu-id="fb264-107">You must reference the Office JavaScript API from inside the `<head>` section of the page to ensure that the API is fully initialized prior to any body elements.</span></span> <span data-ttu-id="fb264-108">Office 应用程序要求外接程序在激活5秒内初始化。</span><span class="sxs-lookup"><span data-stu-id="fb264-108">Office applications require that add-ins initialize within 5 seconds of activation.</span></span> <span data-ttu-id="fb264-109">如果外接程序未在此阈值内激活，则会被声明为无响应，并且用户会看到错误消息。</span><span class="sxs-lookup"><span data-stu-id="fb264-109">If your add-in doesn't activate within this threshold, it will be declared unresponsive and an error message will be displayed to the user.</span></span>

## <a name="api-versioning-and-backward-compatibility"></a><span data-ttu-id="fb264-110">API 版本控制和向后兼容性</span><span class="sxs-lookup"><span data-stu-id="fb264-110">API versioning and backward compatibility</span></span>

<span data-ttu-id="fb264-111">在上面的 HTML 代码段中，CDN URL 中的 " `/1/` 在 `office.js` Office.js 的第1版中指定最新的增量释放。</span><span class="sxs-lookup"><span data-stu-id="fb264-111">In the previous HTML snippet, the `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js.</span></span> <span data-ttu-id="fb264-112">由于 Office JavaScript API 保持向后兼容性，最新版本将继续支持之前在版本1中引入的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="fb264-112">Because the Office JavaScript API maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1.</span></span> <span data-ttu-id="fb264-113">如果需要升级现有项目，请参阅 [更新 Office JAVASCRIPT API 和清单架构文件的版本](update-your-javascript-api-for-office-and-manifest-schema-version.md)。</span><span class="sxs-lookup"><span data-stu-id="fb264-113">If you need to upgrade an existing project, see [Update the version of your Office JavaScript API and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="fb264-p104">如果计划从 AppSource 发布 Office 加载项，必须使用此 CDN 引用。本地引用仅适用于内部、开发和调试应用场景。</span><span class="sxs-lookup"><span data-stu-id="fb264-p104">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!NOTE]
> <span data-ttu-id="fb264-116">要使用预览版 API，请参考 CDN 上的 Office JavaScript API 库预览版：`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`。</span><span class="sxs-lookup"><span data-stu-id="fb264-116">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

## <a name="enabling-intellisense-for-a-typescript-project"></a><span data-ttu-id="fb264-117">为 TypeScript 项目启用 IntelliSense</span><span class="sxs-lookup"><span data-stu-id="fb264-117">Enabling IntelliSense for a TypeScript project</span></span>

<span data-ttu-id="fb264-118">除了参照前面所述的 Office JavaScript API 之外，还可以使用 [jquery.typescript.definitelytyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js)中的类型定义为 TypeScript 加载项项目启用 IntelliSense。</span><span class="sxs-lookup"><span data-stu-id="fb264-118">In addition to referencing the Office JavaScript API as described previously, you can also enable IntelliSense for TypeScript add-in project by using the type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span></span> <span data-ttu-id="fb264-119">若要执行此操作，请在已启用节点的系统提示符 (或 git bash 窗口中) 从项目文件夹的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="fb264-119">To do so, run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder.</span></span> <span data-ttu-id="fb264-120">必须安装 [Node.js](https://nodejs.org)（包括 npm）。</span><span class="sxs-lookup"><span data-stu-id="fb264-120">You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a><span data-ttu-id="fb264-121">预览 Api</span><span class="sxs-lookup"><span data-stu-id="fb264-121">Preview APIs</span></span>

<span data-ttu-id="fb264-122">新的 JavaScript Api 首先在 "预览" 中引入，并在进行充分的测试并需要用户反馈之后成为特定编号的要求集的一部分。</span><span class="sxs-lookup"><span data-stu-id="fb264-122">New JavaScript APIs are first introduced in "preview" and later become part of a specific numbered requirement set after sufficient testing occurs and user feedback is required.</span></span>

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a><span data-ttu-id="fb264-123">另请参阅</span><span class="sxs-lookup"><span data-stu-id="fb264-123">See also</span></span>

- [<span data-ttu-id="fb264-124">了解 Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="fb264-124">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="fb264-125">Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="fb264-125">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)

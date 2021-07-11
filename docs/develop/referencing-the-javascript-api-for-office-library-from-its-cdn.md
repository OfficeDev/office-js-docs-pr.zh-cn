---
title: 参考 Office JavaScript API 库
description: 了解如何在外接程序Office JavaScript API 库和类型定义。
ms.date: 02/18/2021
localization_priority: Normal
ms.openlocfilehash: 04f97412c07cb39f5b2f753c3ce14e56e87c3de5
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349754"
---
# <a name="referencing-the-office-javascript-api-library"></a><span data-ttu-id="6571f-103">参考 Office JavaScript API 库</span><span class="sxs-lookup"><span data-stu-id="6571f-103">Referencing the Office JavaScript API library</span></span>

<span data-ttu-id="6571f-104">Office [JavaScript API](../reference/javascript-api-for-office.md)库提供了外接程序可用于与应用程序应用程序交互Office API。</span><span class="sxs-lookup"><span data-stu-id="6571f-104">The [Office JavaScript API](../reference/javascript-api-for-office.md) library provides the APIs that your add-in can use to interact with the Office application.</span></span> <span data-ttu-id="6571f-105">引用库的最简单方法就是使用内容交付网络 (CDN) HTML 页面的 部分添加 `<script>` `<head>` 以下标记。</span><span class="sxs-lookup"><span data-stu-id="6571f-105">The simplest way to reference the library is to use the content delivery network (CDN) by adding the following `<script>` tag within the `<head>` section of your HTML page.</span></span>

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

<span data-ttu-id="6571f-106">这将在外接程序首次加载时下载并缓存 Office JavaScript API 文件，以确保其对指定版本使用 Office.js 及其关联文件最新的实现。</span><span class="sxs-lookup"><span data-stu-id="6571f-106">This will download and cache the Office JavaScript API files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6571f-107">你必须从页面Office引用 JavaScript API，以确保 API 在任意 body 元素之前 `<head>` 完全初始化。</span><span class="sxs-lookup"><span data-stu-id="6571f-107">You must reference the Office JavaScript API from inside the `<head>` section of the page to ensure that the API is fully initialized prior to any body elements.</span></span>

## <a name="api-versioning-and-backward-compatibility"></a><span data-ttu-id="6571f-108">API 版本控制与向后兼容性</span><span class="sxs-lookup"><span data-stu-id="6571f-108">API versioning and backward compatibility</span></span>

<span data-ttu-id="6571f-109">在前面的 HTML 代码段中，CDN URL 中的 前面的 指定版本 1 中的最新增量 `/1/` `office.js` Office.js。</span><span class="sxs-lookup"><span data-stu-id="6571f-109">In the previous HTML snippet, the `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js.</span></span> <span data-ttu-id="6571f-110">由于 Office JavaScript API 保持向后兼容性，因此最新版本将继续支持版本 1 中之前引入的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="6571f-110">Because the Office JavaScript API maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1.</span></span> <span data-ttu-id="6571f-111">如果需要升级现有项目，请参阅[更新 javaScript API Office清单架构文件的版本](update-your-javascript-api-for-office-and-manifest-schema-version.md)。</span><span class="sxs-lookup"><span data-stu-id="6571f-111">If you need to upgrade an existing project, see [Update the version of your Office JavaScript API and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="6571f-p103">如果计划从 AppSource 发布 Office 加载项，必须使用此 CDN 引用。本地引用仅适用于内部、开发和调试应用场景。</span><span class="sxs-lookup"><span data-stu-id="6571f-p103">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!NOTE]
> <span data-ttu-id="6571f-114">要使用预览版 API，请参考 CDN 上的 Office JavaScript API 库预览版：`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`。</span><span class="sxs-lookup"><span data-stu-id="6571f-114">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

## <a name="enabling-intellisense-for-a-typescript-project"></a><span data-ttu-id="6571f-115">为IntelliSense启用项目</span><span class="sxs-lookup"><span data-stu-id="6571f-115">Enabling IntelliSense for a TypeScript project</span></span>

<span data-ttu-id="6571f-116">除了如前面Office引用 JavaScript API 外，您还可以使用[DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js)中的类型定义为 TypeScript 外接程序项目启用 IntelliSense。</span><span class="sxs-lookup"><span data-stu-id="6571f-116">In addition to referencing the Office JavaScript API as described previously, you can also enable IntelliSense for TypeScript add-in project by using the type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span></span> <span data-ttu-id="6571f-117">为此，请从项目文件夹的根目录 (启用节点的系统提示符或 git bash) 运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="6571f-117">To do so, run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder.</span></span> <span data-ttu-id="6571f-118">必须安装 [Node.js](https://nodejs.org)（包括 npm）。</span><span class="sxs-lookup"><span data-stu-id="6571f-118">You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a><span data-ttu-id="6571f-119">预览 API</span><span class="sxs-lookup"><span data-stu-id="6571f-119">Preview APIs</span></span>

<span data-ttu-id="6571f-120">新的 JavaScript API 首先在"预览版"中引入，之后在经过充分测试且需要用户反馈后成为特定编号要求集的一部分。</span><span class="sxs-lookup"><span data-stu-id="6571f-120">New JavaScript APIs are first introduced in "preview" and later become part of a specific numbered requirement set after sufficient testing occurs and user feedback is required.</span></span>

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a><span data-ttu-id="6571f-121">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6571f-121">See also</span></span>

- [<span data-ttu-id="6571f-122">了解 Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="6571f-122">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="6571f-123">Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="6571f-123">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)

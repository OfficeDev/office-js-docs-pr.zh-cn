---
title: 从内容传送网络 (CDN) 引用适用于 Office 的 JavaScript API 库
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 6b9512d5d0969e185902d7ab9d3227e820c4d0dc
ms.sourcegitcommit: 528577145b2cf0a42bc64c56145d661c4d019fb8
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/02/2019
ms.locfileid: "37353816"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a><span data-ttu-id="d7b70-102">从内容传送网络 (CDN) 引用适用于 Office 的 JavaScript API 库</span><span class="sxs-lookup"><span data-stu-id="d7b70-102">Referencing the JavaScript API for Office library from its content delivery network (CDN)</span></span>

> [!NOTE]
> <span data-ttu-id="d7b70-103">如果想要使用 TypeScript 获取 IntelliSense，除了本文中所述的步骤之外，还需要在项目文件夹根目录的节点支持系统提示框（或 Git Bash 窗口）中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="d7b70-103">In addition to the steps described in this article, if you want to use TypeScript, then to get Intellisense you will need run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder.</span></span> <span data-ttu-id="d7b70-104">必须安装 [Node.js](https://nodejs.org)（包括 npm）。</span><span class="sxs-lookup"><span data-stu-id="d7b70-104">You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>
> 
> ```command&nbsp;line
> npm install --save-dev @types/office-js
> ```

<span data-ttu-id="d7b70-105">[适用于 Office 的 JavaScript API](/office/dev/add-ins/reference/javascript-api-for-office) 库包含 Office.js 文件和关联的主机应用专用 .js 文件（如 Excel-15.js 和 Outlook-15.js）。</span><span class="sxs-lookup"><span data-stu-id="d7b70-105">The [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js.</span></span> 


<span data-ttu-id="d7b70-106">引用 API 的最简单方法是，通过向页面的 `<head>` 标记添加以下 `<script>` 来使用我们的 CDN：</span><span class="sxs-lookup"><span data-stu-id="d7b70-106">The simplest way to reference the API is to use our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

<span data-ttu-id="d7b70-p102">在 CDN URL 中，`office.js` 前面的 `/1/` 指定 Office.js 第 1 版中的最新增量版本。由于适用于 Office 的 JavaScript API 保留向后兼容性，因此最新版本将继续支持之前在第 1 版中引入的 API 成员。如果需要升级现有项目，请参阅[更新适用于 Office 的 JavaScript API 的版本和清单架构文件](update-your-javascript-api-for-office-and-manifest-schema-version.md)。</span><span class="sxs-lookup"><span data-stu-id="d7b70-p102">The  `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js. Because the JavaScript API for Office maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1. If you need to upgrade an existing project, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="d7b70-p103">如果计划从 AppSource 发布 Office 加载项，必须使用此 CDN 引用。本地引用仅适用于内部、开发和调试应用场景。</span><span class="sxs-lookup"><span data-stu-id="d7b70-p103">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d7b70-p104">开发任何 Office 主机应用的加载项时，请从页面的 `<head>` 部分引用适用于 Office 的 JavaScript API。这样可确保 API 先于所有正文元素完全初始化。Office 主机要求，加载项必须在激活后的 5 秒内初始化。如果加载项未在此阈值内激活，则会被声明为无响应，并且用户会看到错误消息。</span><span class="sxs-lookup"><span data-stu-id="d7b70-p104">When you develop an add-in for any Office host application, reference the JavaScript API for Office from inside the `<head>` section of the page. This ensures that the API is fully initialized prior to any body elements. Office hosts require that add-ins initialize within 5 seconds of activation. If your add-in doesn't activate within this threshold, it will be declared unresponsive and an error message will be displayed to the user.</span></span>

## <a name="see-also"></a><span data-ttu-id="d7b70-116">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d7b70-116">See also</span></span>

- [<span data-ttu-id="d7b70-117">了解适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="d7b70-117">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="d7b70-118">适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="d7b70-118">JavaScript API for Office</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)

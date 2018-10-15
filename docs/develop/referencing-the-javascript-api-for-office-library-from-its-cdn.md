---
title: 从内容传送网络 (CDN) 引用适用于 Office 的 JavaScript API 库
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 0ad589ee98342ee72259cddc0957277e9018f186
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505417"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a><span data-ttu-id="bb288-102">从内容传送网络 (CDN) 引用适用于 Office 的 JavaScript API 库</span><span class="sxs-lookup"><span data-stu-id="bb288-102">Referencing the JavaScript API for Office library from its content delivery network (CDN)</span></span>

> [!NOTE]
> <span data-ttu-id="bb288-p101">除了本文中描述的步骤之外，如果要使用 TypeScript，然后获取 Intellisense，您需要在项目文件夹根目录下在启用 Node 的系统提示符（或 git bash 窗口）中运行以下命令。您必须安装 [Node.js](https://nodejs.org)  (其中包括 npm)。</span><span class="sxs-lookup"><span data-stu-id="bb288-p101">In addition to the steps described in this article, if you want to use TypeScript, then to get Intellisense you will need run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder. You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>
> 
> ```
> npm install --save-dev @types/office-js
> ```

<span data-ttu-id="bb288-105">[适用于 Office 的 JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) 库包含 Office.js 文件和关联主机应用的专用 .js 文件（如 Excel-15.js 和 Outlook-15.js）。</span><span class="sxs-lookup"><span data-stu-id="bb288-105">The [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js.</span></span> 


<span data-ttu-id="bb288-106">引用 API 的最简单方法是，通过向页面的 `<head>` 标记添加以下 `<script>` 来使用我们的 CDN：</span><span class="sxs-lookup"><span data-stu-id="bb288-106">The simplest way to reference the API is to use our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

<span data-ttu-id="bb288-p102">在 CDN URL 中，`office.js` 前面的 `/1/` 指定 Office.js 第 1 版中的最新增量版本。由于适用于 Office 的 JavaScript API 保留向后兼容性，因此最新版本将继续支持之前在第 1 版中引入的 API 成员。如果需要升级现有项目，请参阅[更新适用于 Office 的 JavaScript API 的版本和清单架构文件](update-your-javascript-api-for-office-and-manifest-schema-version.md)。</span><span class="sxs-lookup"><span data-stu-id="bb288-p102">The  `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js. Because the JavaScript API for Office maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1. If you need to upgrade an existing project, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="bb288-p103">如果计划从 AppSource 发布 Office 加载项，必须使用此 CDN 引用。本地引用仅适用于内部、开发和调试应用场景。</span><span class="sxs-lookup"><span data-stu-id="bb288-p103">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!IMPORTANT]
>  <span data-ttu-id="bb288-p104">开发任何 Office 主机应用的加载项时，请从页面的 `<head>` 部分引用适用于 Office 的 JavaScript API。这样可确保 API 先于所有正文元素完全初始化。Office 主机要求，加载项必须在激活后的 5 秒内初始化。如果加载项未在此阈值内激活，则会被声明为无响应，并且用户会看到错误消息。</span><span class="sxs-lookup"><span data-stu-id="bb288-p104">When you develop an add-in for any Office host application, reference the JavaScript API for Office from inside the `<head>` section of the page. This ensures that the API is fully initialized prior to any body elements. Office hosts require that add-ins initialize within 5 seconds of activation. If your add-in doesn't activate within this threshold, it will be declared unresponsive and an error message will be displayed to the user.</span></span>       

## <a name="see-also"></a><span data-ttu-id="bb288-116">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bb288-116">See also</span></span>

- [<span data-ttu-id="bb288-117">了解适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="bb288-117">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)    
- [<span data-ttu-id="bb288-118">适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="bb288-118">JavaScript API for Office</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)
    

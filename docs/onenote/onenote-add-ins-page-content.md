---
title: 处理 OneNote 页面内容
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: aef9d80ebb37dacd2c3b5f2ec9d33cb0164d8452
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457612"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="12cb7-102">处理 OneNote 页面内容</span><span class="sxs-lookup"><span data-stu-id="12cb7-102">Work with OneNote page content</span></span> 

<span data-ttu-id="12cb7-103">在 OneNote 外接程序 JavaScript API 中，页面内容由以下对象模型表示。</span><span class="sxs-lookup"><span data-stu-id="12cb7-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![OneNote 页面对象模型图](../images/one-note-om-page.png)

- <span data-ttu-id="12cb7-105">Page 对象包含一组 PageContent 对象。</span><span class="sxs-lookup"><span data-stu-id="12cb7-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="12cb7-106">PageContent 对象包含类型为 Outline、Image 或 Other 的内容。</span><span class="sxs-lookup"><span data-stu-id="12cb7-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="12cb7-107">Outline 对象包含一组 Paragraph 对象。</span><span class="sxs-lookup"><span data-stu-id="12cb7-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="12cb7-108">Paragraph 对象包含 RichText、Image、Table 或 Other 这些内容类型。</span><span class="sxs-lookup"><span data-stu-id="12cb7-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="12cb7-109">若要创建空的 OneNote 页面，请使用下列方法之一：</span><span class="sxs-lookup"><span data-stu-id="12cb7-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="12cb7-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="12cb7-110">Section.addPage</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section#addpage-title-)
- [<span data-ttu-id="12cb7-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="12cb7-111">Page.insertPageAsSibling</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

<span data-ttu-id="12cb7-112">然后使用以下对象中的方法处理页面内容，如 Page.addOutline 和 Outline.appendHtml。</span><span class="sxs-lookup"><span data-stu-id="12cb7-112">Then use methods in the following objects to work with the page content, such as Page.addOutline and Outline.appendHtml.</span></span> 

- [<span data-ttu-id="12cb7-113">Page</span><span class="sxs-lookup"><span data-stu-id="12cb7-113">Page</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.page)
- [<span data-ttu-id="12cb7-114">Outline</span><span class="sxs-lookup"><span data-stu-id="12cb7-114">Outline</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.outline)
- [<span data-ttu-id="12cb7-115">Paragraph</span><span class="sxs-lookup"><span data-stu-id="12cb7-115">Paragraph</span></span>](https://docs.microsoft.com/javascript/api/onenote/onenote.paragraph)

<span data-ttu-id="12cb7-p101">OneNote 页面的内容和结构由 HTML 进行表示。只有一部分 HTML 可用于创建或更新页面内容，如下所述。</span><span class="sxs-lookup"><span data-stu-id="12cb7-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="12cb7-118">受支持的 HTML</span><span class="sxs-lookup"><span data-stu-id="12cb7-118">Supported HTML</span></span>

<span data-ttu-id="12cb7-119">OneNote 外接程序 JavaScript API 支持使用以下 HTML 创建和更新页面内容：</span><span class="sxs-lookup"><span data-stu-id="12cb7-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="12cb7-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="12cb7-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span></span> 
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="12cb7-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="12cb7-121">`<ul>`, `<ol>`, `<li>`</span></span> 
- <span data-ttu-id="12cb7-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="12cb7-122">`<table>`, `<tr>`, `<td>`</span></span>
- <span data-ttu-id="12cb7-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="12cb7-123">`<h1>` ... `<h6>`</span></span>
- <span data-ttu-id="12cb7-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="12cb7-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span></span>

> [!NOTE]
> <span data-ttu-id="12cb7-125">将 HTML 导入 OneNote 合并空白。</span><span class="sxs-lookup"><span data-stu-id="12cb7-125">Importing HTML into OneNote consolidates whitespace.</span></span> <span data-ttu-id="12cb7-126">生成的内容将粘贴到一个大纲中。</span><span class="sxs-lookup"><span data-stu-id="12cb7-126">The resulting content is pasted into one outline.</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="12cb7-127">访问页面内容</span><span class="sxs-lookup"><span data-stu-id="12cb7-127">Accessing page contents</span></span>

<span data-ttu-id="12cb7-p103">只可通过 `Page#load` 访问当前活动页的*页面内容*。若要更改活动页，请调用 `navigateToPage($page)`。</span><span class="sxs-lookup"><span data-stu-id="12cb7-p103">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="12cb7-130">仍可查询任何页面的元数据（如标题）。</span><span class="sxs-lookup"><span data-stu-id="12cb7-130">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="12cb7-131">另请参阅</span><span class="sxs-lookup"><span data-stu-id="12cb7-131">See also</span></span>

- [<span data-ttu-id="12cb7-132">OneNote JavaScript API 编程概述</span><span class="sxs-lookup"><span data-stu-id="12cb7-132">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="12cb7-133">OneNote JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="12cb7-133">OneNote JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="12cb7-134">Rubric Grader 示例</span><span class="sxs-lookup"><span data-stu-id="12cb7-134">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="12cb7-135">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="12cb7-135">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

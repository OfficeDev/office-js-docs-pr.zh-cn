---
title: 处理 OneNote 页面内容
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: f60cdee7eb549acc0f2c84a1aa9acea7fe77274a
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872184"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="e4db3-102">处理 OneNote 页面内容</span><span class="sxs-lookup"><span data-stu-id="e4db3-102">Work with OneNote page content</span></span>

<span data-ttu-id="e4db3-103">在 OneNote 外接程序 JavaScript API 中，页面内容由以下对象模型表示。</span><span class="sxs-lookup"><span data-stu-id="e4db3-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![OneNote 页面对象模型图](../images/one-note-om-page.png)

- <span data-ttu-id="e4db3-105">Page 对象包含一组 PageContent 对象。</span><span class="sxs-lookup"><span data-stu-id="e4db3-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="e4db3-106">PageContent 对象包含类型为 Outline、Image 或 Other 的内容。</span><span class="sxs-lookup"><span data-stu-id="e4db3-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="e4db3-107">Outline 对象包含一组 Paragraph 对象。</span><span class="sxs-lookup"><span data-stu-id="e4db3-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="e4db3-108">Paragraph 对象包含 RichText、Image、Table 或 Other 这些内容类型。</span><span class="sxs-lookup"><span data-stu-id="e4db3-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="e4db3-109">若要创建空的 OneNote 页面，请使用下列方法之一：</span><span class="sxs-lookup"><span data-stu-id="e4db3-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="e4db3-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="e4db3-110">Section.addPage</span></span>](/javascript/api/onenote/onenote.section#addpage-title-)
- [<span data-ttu-id="e4db3-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="e4db3-111">Page.insertPageAsSibling</span></span>](/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

<span data-ttu-id="e4db3-112">然后使用以下对象中的方法处理页面内容，如 `Page.addOutline` 和 `Outline.appendHtml`。</span><span class="sxs-lookup"><span data-stu-id="e4db3-112">Then use methods in the following objects to work with the page content, such as `Page.addOutline` and `Outline.appendHtml`.</span></span>

- [<span data-ttu-id="e4db3-113">Page</span><span class="sxs-lookup"><span data-stu-id="e4db3-113">Page</span></span>](/javascript/api/onenote/onenote.page)
- [<span data-ttu-id="e4db3-114">Outline</span><span class="sxs-lookup"><span data-stu-id="e4db3-114">Outline</span></span>](/javascript/api/onenote/onenote.outline)
- [<span data-ttu-id="e4db3-115">Paragraph</span><span class="sxs-lookup"><span data-stu-id="e4db3-115">Paragraph</span></span>](/javascript/api/onenote/onenote.paragraph)

<span data-ttu-id="e4db3-p101">OneNote 页面的内容和结构由 HTML 进行表示。只有一部分 HTML 可用于创建或更新页面内容，如下所述。</span><span class="sxs-lookup"><span data-stu-id="e4db3-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="e4db3-118">受支持的 HTML</span><span class="sxs-lookup"><span data-stu-id="e4db3-118">Supported HTML</span></span>

<span data-ttu-id="e4db3-119">OneNote 外接程序 JavaScript API 支持使用以下 HTML 创建和更新页面内容：</span><span class="sxs-lookup"><span data-stu-id="e4db3-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="e4db3-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="e4db3-120"></span></span>
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="e4db3-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="e4db3-121"></span></span>
- <span data-ttu-id="e4db3-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="e4db3-122"></span></span>
- <span data-ttu-id="e4db3-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="e4db3-123"></span></span>
- <span data-ttu-id="e4db3-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="e4db3-124"></span></span>

> [!NOTE]
> <span data-ttu-id="e4db3-125">将 HTML 导入 OneNote 合并空白。</span><span class="sxs-lookup"><span data-stu-id="e4db3-125">Importing HTML into OneNote consolidates whitespace.</span></span> <span data-ttu-id="e4db3-126">生成的内容将粘贴到一个大纲中。</span><span class="sxs-lookup"><span data-stu-id="e4db3-126">The resulting content is pasted into one outline.</span></span>

<span data-ttu-id="e4db3-127">OneNote 会尽力将 HTML 翻译成页面内容，同时确保用户的安全性。</span><span class="sxs-lookup"><span data-stu-id="e4db3-127">OneNote does its best to translate HTML into page content while ensuring security for users.</span></span> <span data-ttu-id="e4db3-128">HTML 和 CSS 标准并不完全与 OneNote 的内容模型匹配，因此，会存在外观上的差异，尤其是采用 CSS 样式时。</span><span class="sxs-lookup"><span data-stu-id="e4db3-128">HTML and CSS standards do not exactly match OneNote's content model, so there will be differences in appearances, particularly with CSS stylings.</span></span> <span data-ttu-id="e4db3-129">如果需要特定格式，则建议使用 JavaScript 对象。</span><span class="sxs-lookup"><span data-stu-id="e4db3-129">We recommend using the JavaScript objects if specific formatting is needed.</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="e4db3-130">访问页面内容</span><span class="sxs-lookup"><span data-stu-id="e4db3-130">Accessing page contents</span></span>

<span data-ttu-id="e4db3-p104">只可通过 `Page#load` 访问当前活动页的*页面内容*。若要更改活动页，请调用 `navigateToPage($page)`。</span><span class="sxs-lookup"><span data-stu-id="e4db3-p104">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="e4db3-133">仍可查询任何页面的元数据（如标题）。</span><span class="sxs-lookup"><span data-stu-id="e4db3-133">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="e4db3-134">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e4db3-134">See also</span></span>

- [<span data-ttu-id="e4db3-135">OneNote JavaScript API 编程概述</span><span class="sxs-lookup"><span data-stu-id="e4db3-135">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="e4db3-136">OneNote JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="e4db3-136">OneNote JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="e4db3-137">Rubric Grader 示例</span><span class="sxs-lookup"><span data-stu-id="e4db3-137">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="e4db3-138">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="e4db3-138">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

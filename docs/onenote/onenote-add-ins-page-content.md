---
title: 处理 OneNote 页面内容
description: 了解如何使用 JavaScript API OneNote页面内容。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: f506617bfdbc97e94f8fb16930dfc2a935385d5f
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349047"
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="7a1d8-103">处理 OneNote 页面内容</span><span class="sxs-lookup"><span data-stu-id="7a1d8-103">Work with OneNote page content</span></span>

<span data-ttu-id="7a1d8-104">在 OneNote 外接程序 JavaScript API 中，页面内容由以下对象模型表示。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-104">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![OneNote页面对象模型图。](../images/one-note-om-page.png)

- <span data-ttu-id="7a1d8-106">Page 对象包含一组 PageContent 对象。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-106">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="7a1d8-107">PageContent 对象包含类型为 Outline、Image 或 Other 的内容。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-107">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="7a1d8-108">Outline 对象包含一组 Paragraph 对象。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-108">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="7a1d8-109">Paragraph 对象包含 RichText、Image、Table 或 Other 这些内容类型。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-109">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="7a1d8-110">若要创建空OneNote页，请使用下列方法之一。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-110">To create an empty OneNote page, use one of the following methods.</span></span>

- [<span data-ttu-id="7a1d8-111">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="7a1d8-111">Section.addPage</span></span>](/javascript/api/onenote/onenote.section#addpage-title-)
- [<span data-ttu-id="7a1d8-112">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="7a1d8-112">Page.insertPageAsSibling</span></span>](/javascript/api/onenote/onenote.section#insertsectionassibling-location--title-)

<span data-ttu-id="7a1d8-113">然后使用以下对象中的方法处理页面内容，如 `Page.addOutline` 和 `Outline.appendHtml`。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-113">Then use methods in the following objects to work with the page content, such as `Page.addOutline` and `Outline.appendHtml`.</span></span>

- [<span data-ttu-id="7a1d8-114">Page</span><span class="sxs-lookup"><span data-stu-id="7a1d8-114">Page</span></span>](/javascript/api/onenote/onenote.page)
- [<span data-ttu-id="7a1d8-115">Outline</span><span class="sxs-lookup"><span data-stu-id="7a1d8-115">Outline</span></span>](/javascript/api/onenote/onenote.outline)
- [<span data-ttu-id="7a1d8-116">Paragraph</span><span class="sxs-lookup"><span data-stu-id="7a1d8-116">Paragraph</span></span>](/javascript/api/onenote/onenote.paragraph)

<span data-ttu-id="7a1d8-p101">OneNote 页面的内容和结构由 HTML 进行表示。只有一部分 HTML 可用于创建或更新页面内容，如下所述。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="7a1d8-119">受支持的 HTML</span><span class="sxs-lookup"><span data-stu-id="7a1d8-119">Supported HTML</span></span>

<span data-ttu-id="7a1d8-120">加载项OneNote JavaScript API 支持以下 HTML 来创建和更新页面内容。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-120">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content.</span></span>

- <span data-ttu-id="7a1d8-121">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="7a1d8-121">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span></span>
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="7a1d8-122">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="7a1d8-122">`<ul>`, `<ol>`, `<li>`</span></span>
- <span data-ttu-id="7a1d8-123">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="7a1d8-123">`<table>`, `<tr>`, `<td>`</span></span>
- <span data-ttu-id="7a1d8-124">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="7a1d8-124">`<h1>` ... `<h6>`</span></span>
- <span data-ttu-id="7a1d8-125">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="7a1d8-125">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span></span>

> [!NOTE]
> <span data-ttu-id="7a1d8-126">将 HTML 导入 OneNote 合并空白。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-126">Importing HTML into OneNote consolidates whitespace.</span></span> <span data-ttu-id="7a1d8-127">生成的内容将粘贴到一个大纲中。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-127">The resulting content is pasted into one outline.</span></span>

<span data-ttu-id="7a1d8-128">OneNote 会尽力将 HTML 翻译成页面内容，同时确保用户的安全性。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-128">OneNote does its best to translate HTML into page content while ensuring security for users.</span></span> <span data-ttu-id="7a1d8-129">HTML 和 CSS 标准并不完全与 OneNote 的内容模型匹配，因此，会存在外观上的差异，尤其是采用 CSS 样式时。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-129">HTML and CSS standards do not exactly match OneNote's content model, so there will be differences in appearances, particularly with CSS stylings.</span></span> <span data-ttu-id="7a1d8-130">如果需要特定格式，则建议使用 JavaScript 对象。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-130">We recommend using the JavaScript objects if specific formatting is needed.</span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="7a1d8-131">访问页面内容</span><span class="sxs-lookup"><span data-stu-id="7a1d8-131">Accessing page contents</span></span>

<span data-ttu-id="7a1d8-p104">只可通过 `Page#load` 访问当前活动页的 *页面内容*。若要更改活动页，请调用 `navigateToPage($page)`。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-p104">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="7a1d8-134">仍可查询任何页面的元数据（如标题）。</span><span class="sxs-lookup"><span data-stu-id="7a1d8-134">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="7a1d8-135">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7a1d8-135">See also</span></span>

- [<span data-ttu-id="7a1d8-136">OneNote JavaScript API 编程概述</span><span class="sxs-lookup"><span data-stu-id="7a1d8-136">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="7a1d8-137">OneNote JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="7a1d8-137">OneNote JavaScript API reference</span></span>](../reference/overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="7a1d8-138">Rubric Grader 示例</span><span class="sxs-lookup"><span data-stu-id="7a1d8-138">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="7a1d8-139">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="7a1d8-139">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

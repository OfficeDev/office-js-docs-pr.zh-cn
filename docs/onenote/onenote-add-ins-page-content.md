---
title: ?? OneNote ????
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d05f251a798a7670983187bfa4c80140b30f6147
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="work-with-onenote-page-content"></a><span data-ttu-id="67ca9-102">?? OneNote ????</span><span class="sxs-lookup"><span data-stu-id="67ca9-102">Work with OneNote page content</span></span> 

<span data-ttu-id="67ca9-103">? OneNote ???? JavaScript API ????????????????</span><span class="sxs-lookup"><span data-stu-id="67ca9-103">In the OneNote add-ins JavaScript API, page content is represented by the following object model.</span></span>

  ![OneNote ???????](../images/one-note-om-page.png)

- <span data-ttu-id="67ca9-105">Page ?????? PageContent ???</span><span class="sxs-lookup"><span data-stu-id="67ca9-105">A Page object contains a collection of PageContent objects.</span></span>
- <span data-ttu-id="67ca9-106">PageContent ??????? Outline?Image ? Other ????</span><span class="sxs-lookup"><span data-stu-id="67ca9-106">A PageContent object contains a content type of Outline, Image, or Other.</span></span>
- <span data-ttu-id="67ca9-107">Outline ?????? Paragraph ???</span><span class="sxs-lookup"><span data-stu-id="67ca9-107">An Outline object contains a collection of Paragraph objects.</span></span>
- <span data-ttu-id="67ca9-108">Paragraph ???? RichText?Image?Table ? Other ???????</span><span class="sxs-lookup"><span data-stu-id="67ca9-108">A Paragraph object contains a content type of RichText, Image, Table, or Other.</span></span>

<span data-ttu-id="67ca9-109">?????? OneNote ?????????????</span><span class="sxs-lookup"><span data-stu-id="67ca9-109">To create an empty OneNote page, use one of the following methods:</span></span>

- [<span data-ttu-id="67ca9-110">Section.addPage</span><span class="sxs-lookup"><span data-stu-id="67ca9-110">Section.addPage</span></span>](https://dev.office.com/reference/add-ins/onenote/section#addpagetitle-string)
- [<span data-ttu-id="67ca9-111">Page.insertPageAsSibling</span><span class="sxs-lookup"><span data-stu-id="67ca9-111">Page.insertPageAsSibling</span></span>](https://dev.office.com/reference/add-ins/onenote/page#insertpageassiblinglocation-string-title-string)

<span data-ttu-id="67ca9-112">???????????????????? Page.addOutline ? Outline.appendHtml?</span><span class="sxs-lookup"><span data-stu-id="67ca9-112">Then use methods in the following objects to work with the page content, such as Page.addOutline and Outline.appendHtml.</span></span> 

- [<span data-ttu-id="67ca9-113">Page</span><span class="sxs-lookup"><span data-stu-id="67ca9-113">Page</span></span>](https://dev.office.com/reference/add-ins/onenote/page)
- [<span data-ttu-id="67ca9-114">Outline</span><span class="sxs-lookup"><span data-stu-id="67ca9-114">Outline</span></span>](https://dev.office.com/reference/add-ins/onenote/outline)
- [<span data-ttu-id="67ca9-115">Paragraph</span><span class="sxs-lookup"><span data-stu-id="67ca9-115">Paragraph</span></span>](https://dev.office.com/reference/add-ins/onenote/paragraph)

<span data-ttu-id="67ca9-p101">OneNote ????????? HTML ?????????? HTML ??????????????????</span><span class="sxs-lookup"><span data-stu-id="67ca9-p101">The content and structure of a OneNote page are represented by HTML. Only a subset of HTML is supported for creating or updating page content, as described below.</span></span>

## <a name="supported-html"></a><span data-ttu-id="67ca9-118">???? HTML</span><span class="sxs-lookup"><span data-stu-id="67ca9-118">Supported HTML</span></span>

<span data-ttu-id="67ca9-119">OneNote ???? JavaScript API ?????? HTML ??????????</span><span class="sxs-lookup"><span data-stu-id="67ca9-119">The OneNote add-in JavaScript API supports the following HTML for creating and updating page content:</span></span>

- <span data-ttu-id="67ca9-120">`<html>`, `<body>`, `<div>`, `<span>`, `<br/>`</span><span class="sxs-lookup"><span data-stu-id="67ca9-120"></span></span> 
- `<p>`
- `<img>`
- `<a>`
- <span data-ttu-id="67ca9-121">`<ul>`, `<ol>`, `<li>`</span><span class="sxs-lookup"><span data-stu-id="67ca9-121"></span></span> 
- <span data-ttu-id="67ca9-122">`<table>`, `<tr>`, `<td>`</span><span class="sxs-lookup"><span data-stu-id="67ca9-122"></span></span>
- <span data-ttu-id="67ca9-123">`<h1>` ... `<h6>`</span><span class="sxs-lookup"><span data-stu-id="67ca9-123"></span></span>
- <span data-ttu-id="67ca9-124">`<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`</span><span class="sxs-lookup"><span data-stu-id="67ca9-124"></span></span>

## <a name="accessing-page-contents"></a><span data-ttu-id="67ca9-125">??????</span><span class="sxs-lookup"><span data-stu-id="67ca9-125">Accessing page contents</span></span>

<span data-ttu-id="67ca9-p102">???? `Page#load` ????????*????*???????????? `navigateToPage($page)`?</span><span class="sxs-lookup"><span data-stu-id="67ca9-p102">You are only able to access *Page Content* via `Page#load` for the currently active page. To change the active  page, invoke `navigateToPage($page)`.</span></span>

<span data-ttu-id="67ca9-128">??????????????????</span><span class="sxs-lookup"><span data-stu-id="67ca9-128">Metadata such as title can still be queried for any page.</span></span>

## <a name="see-also"></a><span data-ttu-id="67ca9-129">????</span><span class="sxs-lookup"><span data-stu-id="67ca9-129">See also</span></span>

- [<span data-ttu-id="67ca9-130">OneNote JavaScript API ????</span><span class="sxs-lookup"><span data-stu-id="67ca9-130">OneNote JavaScript API programming overview</span></span>](onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="67ca9-131">OneNote JavaScript API ??</span><span class="sxs-lookup"><span data-stu-id="67ca9-131">OneNote JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="67ca9-132">Rubric Grader ??</span><span class="sxs-lookup"><span data-stu-id="67ca9-132">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="67ca9-133">Office ???????</span><span class="sxs-lookup"><span data-stu-id="67ca9-133">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

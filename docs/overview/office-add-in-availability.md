---
title: Office 客户端应用程序和平台的 Office 加载项可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 10/14/2020
localization_priority: Priority
ms.openlocfilehash: 8cd6ccf45cd63a99f70155035d4062e22fa1cbcf
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626559"
---
# <a name="office-client-application-and-platform-availability-for-office-add-ins"></a><span data-ttu-id="fea8c-103">Office 客户端应用程序和平台的 Office 加载项可用性</span><span class="sxs-lookup"><span data-stu-id="fea8c-103">Office client application and platform availability for Office Add-ins</span></span>

<span data-ttu-id="fea8c-p101">为了能够按预期运行，Office 加载项可能会依赖特定的 Office 应用程序、要求集、API 成员或 API 版本。下表列出了每个 Office 应用程序目前所支持的平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="fea8c-p101">To work as expected, your Office Add-in might depend on a specific Office application, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="fea8c-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="fea8c-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="fea8c-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="fea8c-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="fea8c-108">Excel</span><span class="sxs-lookup"><span data-stu-id="fea8c-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="fea8c-109">平台</span><span class="sxs-lookup"><span data-stu-id="fea8c-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="fea8c-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="fea8c-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="fea8c-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="fea8c-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="fea8c-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="fea8c-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="fea8c-113">Office on the web</span></span></td>
    <td><span data-ttu-id="fea8c-114">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-114">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-115">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-115">
      - Content</span></span><br><span data-ttu-id="fea8c-116">
      - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="fea8c-116">
      - Custom Functions</span></span><br><span data-ttu-id="fea8c-117">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-117">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-118">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-118">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fea8c-119">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-119">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fea8c-120">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-120">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fea8c-121">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-121">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fea8c-122">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-122">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fea8c-123">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-123">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fea8c-124">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-124">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fea8c-125">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-125">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fea8c-126">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-126">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="fea8c-127">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-127">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="fea8c-128">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-128">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="fea8c-129">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-129">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="fea8c-130">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-130">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-131">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-131">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-132">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-132">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-133">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-133">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-134">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-134">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-135">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-135">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-136">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-136">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-137">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-137">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-138">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-138">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-139">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-139">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-140">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-140">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-141">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-141">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-142">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-142">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-143">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="fea8c-143">Office on Windows</span></span><br><span data-ttu-id="fea8c-144">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="fea8c-144">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="fea8c-145">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-145">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-146">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-146">
      - Content</span></span><br><span data-ttu-id="fea8c-147">
      - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="fea8c-147">
      - Custom Functions</span></span><br><span data-ttu-id="fea8c-148">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-148">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-149">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-149">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fea8c-150">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-150">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fea8c-151">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-151">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fea8c-152">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-152">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fea8c-153">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-153">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fea8c-154">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-154">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fea8c-155">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-155">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fea8c-156">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-156">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fea8c-157">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-157">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="fea8c-158">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-158">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="fea8c-159">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-159">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="fea8c-160">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-160">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-161">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-161">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fea8c-162">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-162">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-163">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-163">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-164">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-164">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-165">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-165">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-166">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-166">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-167">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-167">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-168">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-168">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-169">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-169">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-170">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-170">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-171">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-171">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-172">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-172">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-173">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-173">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-174">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-174">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-175">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="fea8c-175">Office 2019 on Windows</span></span><br><span data-ttu-id="fea8c-176">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-177">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-177">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-178">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-178">
      - Content</span></span><br><span data-ttu-id="fea8c-179">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-179">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-180">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-180">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fea8c-181">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-181">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fea8c-182">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-182">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fea8c-183">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-183">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fea8c-184">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-184">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fea8c-185">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-185">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fea8c-186">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-186">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fea8c-187">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-187">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fea8c-188">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-188">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-189">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-189">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-190">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-190">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-191">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-191">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-192">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-192">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-193">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-193">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-194">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-194">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-195">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-195">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-196">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-196">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-197">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-197">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-198">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-198">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-199">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-199">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-200">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-200">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-201">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-201">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-202">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="fea8c-202">Office 2016 on Windows</span></span><br><span data-ttu-id="fea8c-203">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-204">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-204">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-205">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-205">
      - Content</span></span> </td>
    <td><span data-ttu-id="fea8c-206">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-206">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fea8c-207">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fea8c-207">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="fea8c-208">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-208">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-209">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-209">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-210">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-210">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-211">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-211">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-212">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-212">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-213">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-213">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-214">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-214">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-215">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-215">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-216">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-216">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-217">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-217">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-218">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-218">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-219">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-219">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-220">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-220">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-221">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="fea8c-221">Office 2013 on Windows</span></span><br><span data-ttu-id="fea8c-222">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-223">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-223">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-224">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-224">
      - Content</span></span> </td>
    <td><span data-ttu-id="fea8c-225">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fea8c-225">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="fea8c-226">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-226">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-227">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-227">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-228">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-228">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-229">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-229">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-230">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-230">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-231">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-231">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-232">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-232">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-233">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-233">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-234">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-234">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-235">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-235">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-236">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-236">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-237">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-237">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-238">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="fea8c-238">Office on iPad</span></span><br><span data-ttu-id="fea8c-239">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="fea8c-239">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="fea8c-240">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-240">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-241">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-241">
      - Content</span></span> </td>
    <td><span data-ttu-id="fea8c-242">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-242">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fea8c-243">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-243">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fea8c-244">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-244">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fea8c-245">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-245">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fea8c-246">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-246">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fea8c-247">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-247">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fea8c-248">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-248">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fea8c-249">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-249">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fea8c-250">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-250">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="fea8c-251">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-251">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="fea8c-252">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-252">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="fea8c-253">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-253">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-254">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-254">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-255">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-255">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-256">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-256">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-257">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-257">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-258">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-258">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-259">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-259">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-260">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-260">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-261">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-261">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-262">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-262">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-263">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-263">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-264">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-264">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-265">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-265">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-266">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="fea8c-266">Office on Mac</span></span><br><span data-ttu-id="fea8c-267">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="fea8c-267">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="fea8c-268">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-268">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-269">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-269">
      - Content</span></span><br><span data-ttu-id="fea8c-270">
      - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="fea8c-270">
      - Custom Functions</span></span><br><span data-ttu-id="fea8c-271">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-271">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-272">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-272">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fea8c-273">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-273">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fea8c-274">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-274">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fea8c-275">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-275">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fea8c-276">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-276">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fea8c-277">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-277">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fea8c-278">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-278">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fea8c-279">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-279">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fea8c-280">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-280">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="fea8c-281">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-281">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="fea8c-282">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-282">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="fea8c-283">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-283">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-284">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-284">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fea8c-285">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-285">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-286">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-286">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-287">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-287">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-288">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-288">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-289">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-289">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-290">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-290">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-291">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-291">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-292">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-292">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-293">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-293">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-294">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-294">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-295">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-295">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-296">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-296">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-297">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-297">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-298">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-298">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-299">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="fea8c-299">Office 2019 on Mac</span></span><br><span data-ttu-id="fea8c-300">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-300">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-301">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-301">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-302">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-302">
      - Content</span></span><br><span data-ttu-id="fea8c-303">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-303">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-304">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-304">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fea8c-305">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-305">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fea8c-306">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-306">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fea8c-307">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-307">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fea8c-308">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-308">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fea8c-309">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-309">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fea8c-310">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-310">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fea8c-311">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-311">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fea8c-312">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-312">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-313">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-313">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-314">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-314">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-315">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-315">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-316">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-316">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-317">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-317">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-318">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-318">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-319">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-319">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-320">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-320">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-321">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-321">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-322">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-322">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-323">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-323">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-324">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-324">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-325">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-325">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-326">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-326">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-327">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="fea8c-327">Office 2016 on Mac</span></span><br><span data-ttu-id="fea8c-328">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-329">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-329">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-330">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-330">
      - Content</span></span> </td>
    <td><span data-ttu-id="fea8c-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fea8c-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fea8c-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="fea8c-333">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-333">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-334">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-334">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-335">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-335">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-336">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-336">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-337">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-337">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-338">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-338">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-339">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-339">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-340">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-340">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-341">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-341">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-342">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-342">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-345">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-345">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-346">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-346">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="fea8c-347">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="fea8c-347">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="fea8c-348">自定义函数（仅 Excel）</span><span class="sxs-lookup"><span data-stu-id="fea8c-348">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="fea8c-349">平台</span><span class="sxs-lookup"><span data-stu-id="fea8c-349">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="fea8c-350">扩展点</span><span class="sxs-lookup"><span data-stu-id="fea8c-350">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="fea8c-351">API 要求集</span><span class="sxs-lookup"><span data-stu-id="fea8c-351">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="fea8c-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="fea8c-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-353">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="fea8c-353">Office on the web</span></span></td>
    <td><span data-ttu-id="fea8c-354">- 自定义函数</span><span class="sxs-lookup"><span data-stu-id="fea8c-354">- Custom Functions</span></span></td>
    <td><span data-ttu-id="fea8c-355">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-355">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-356">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="fea8c-356">Office on Windows</span></span><br><span data-ttu-id="fea8c-357">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="fea8c-357">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="fea8c-358">- 自定义函数</span><span class="sxs-lookup"><span data-stu-id="fea8c-358">- Custom Functions</span></span></td>
    <td><span data-ttu-id="fea8c-359">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-359">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-360">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="fea8c-360">Office on Mac</span></span><br><span data-ttu-id="fea8c-361">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="fea8c-361">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="fea8c-362">- 自定义函数</span><span class="sxs-lookup"><span data-stu-id="fea8c-362">- Custom Functions</span></span></td>
    <td><span data-ttu-id="fea8c-363">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-363">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td></td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="fea8c-364">Outlook</span><span class="sxs-lookup"><span data-stu-id="fea8c-364">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fea8c-365">平台</span><span class="sxs-lookup"><span data-stu-id="fea8c-365">Platform</span></span></th>
    <th><span data-ttu-id="fea8c-366">扩展点</span><span class="sxs-lookup"><span data-stu-id="fea8c-366">Extension points</span></span></th>
    <th><span data-ttu-id="fea8c-367">API 要求集</span><span class="sxs-lookup"><span data-stu-id="fea8c-367">API requirement sets</span></span></th>
    <th><span data-ttu-id="fea8c-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="fea8c-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-369">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="fea8c-369">Office on the web</span></span><br><span data-ttu-id="fea8c-370">（新式）</span><span class="sxs-lookup"><span data-stu-id="fea8c-370">(modern)</span></span></td>
    <td><span data-ttu-id="fea8c-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fea8c-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fea8c-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fea8c-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fea8c-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fea8c-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fea8c-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fea8c-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fea8c-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fea8c-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="fea8c-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="fea8c-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span><br><span data-ttu-id="fea8c-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.9">Mailbox 1.9</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.9">Mailbox 1.9</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-385">不可用</span><span class="sxs-lookup"><span data-stu-id="fea8c-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-386">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="fea8c-386">Office on the web</span></span><br><span data-ttu-id="fea8c-387">（经典）</span><span class="sxs-lookup"><span data-stu-id="fea8c-387">(classic)</span></span></td>
    <td><span data-ttu-id="fea8c-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fea8c-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fea8c-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fea8c-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fea8c-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fea8c-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fea8c-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fea8c-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fea8c-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fea8c-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-399">不可用</span><span class="sxs-lookup"><span data-stu-id="fea8c-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-400">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="fea8c-400">Office on Windows</span></span><br><span data-ttu-id="fea8c-401">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="fea8c-401">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="fea8c-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fea8c-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fea8c-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fea8c-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fea8c-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="fea8c-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fea8c-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fea8c-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fea8c-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fea8c-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fea8c-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="fea8c-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="fea8c-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span><br><span data-ttu-id="fea8c-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9">Mailbox 1.9</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9">Mailbox 1.9</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-417">不可用</span><span class="sxs-lookup"><span data-stu-id="fea8c-417">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-418">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="fea8c-418">Office 2019 on Windows</span></span><br><span data-ttu-id="fea8c-419">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-419">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fea8c-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fea8c-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fea8c-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fea8c-424">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-424">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="fea8c-425">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-425">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fea8c-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fea8c-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fea8c-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fea8c-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fea8c-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="fea8c-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-433">不可用</span><span class="sxs-lookup"><span data-stu-id="fea8c-433">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-434">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="fea8c-434">Office 2016 on Windows</span></span><br><span data-ttu-id="fea8c-435">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-435">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fea8c-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fea8c-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fea8c-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fea8c-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="fea8c-441">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-441">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fea8c-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fea8c-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fea8c-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*
    </span></span></td>
    <td><span data-ttu-id="fea8c-446">不可用</span><span class="sxs-lookup"><span data-stu-id="fea8c-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-447">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="fea8c-447">Office 2013 on Windows</span></span><br><span data-ttu-id="fea8c-448">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-448">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fea8c-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fea8c-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fea8c-452">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-452">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    </td>
    <td><span data-ttu-id="fea8c-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fea8c-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fea8c-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="fea8c-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="fea8c-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*
    </span></span></td>
    <td><span data-ttu-id="fea8c-457">不可用</span><span class="sxs-lookup"><span data-stu-id="fea8c-457">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-458">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="fea8c-458">Office on iOS</span></span><br><span data-ttu-id="fea8c-459">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="fea8c-459">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="fea8c-460">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-460">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fea8c-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fea8c-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fea8c-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fea8c-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fea8c-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-467">不可用</span><span class="sxs-lookup"><span data-stu-id="fea8c-467">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-468">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="fea8c-468">Office on Mac</span></span><br><span data-ttu-id="fea8c-469">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="fea8c-469">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="fea8c-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fea8c-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fea8c-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fea8c-473">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-473">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fea8c-474">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-474">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fea8c-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fea8c-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fea8c-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fea8c-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fea8c-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="fea8c-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="fea8c-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-483">不可用</span><span class="sxs-lookup"><span data-stu-id="fea8c-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-484">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="fea8c-484">Office 2019 on Mac</span></span><br><span data-ttu-id="fea8c-485">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-485">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fea8c-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fea8c-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fea8c-489">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-489">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fea8c-490">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-490">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fea8c-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fea8c-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fea8c-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fea8c-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fea8c-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-497">不可用</span><span class="sxs-lookup"><span data-stu-id="fea8c-497">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-498">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="fea8c-498">Office 2016 on Mac</span></span><br><span data-ttu-id="fea8c-499">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-499">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fea8c-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="fea8c-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="fea8c-503">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-503">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="fea8c-504">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-504">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fea8c-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fea8c-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fea8c-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fea8c-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fea8c-510">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-510">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-511">不可用</span><span class="sxs-lookup"><span data-stu-id="fea8c-511">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-512">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="fea8c-512">Office on Android</span></span><br><span data-ttu-id="fea8c-513">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="fea8c-513">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="fea8c-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="fea8c-515">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">约会组织者（撰写）：联机会议</a> （预览）</span><span class="sxs-lookup"><span data-stu-id="fea8c-515">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="fea8c-516">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-516">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fea8c-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fea8c-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fea8c-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fea8c-521">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-521">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-522">不可用</span><span class="sxs-lookup"><span data-stu-id="fea8c-522">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="fea8c-523">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="fea8c-523">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fea8c-524">要求集的客户端支持可能受到 Exchange 服务器支持的限制。</span><span class="sxs-lookup"><span data-stu-id="fea8c-524">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="fea8c-525">有关 Exchange 服务器和 Outlook 客户端支持的要求集范围的详细信息，请参阅 [Outlook JavaScript API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="fea8c-525">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="fea8c-526">Word</span><span class="sxs-lookup"><span data-stu-id="fea8c-526">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fea8c-527">平台</span><span class="sxs-lookup"><span data-stu-id="fea8c-527">Platform</span></span></th>
    <th><span data-ttu-id="fea8c-528">扩展点</span><span class="sxs-lookup"><span data-stu-id="fea8c-528">Extension points</span></span></th>
    <th><span data-ttu-id="fea8c-529">API 要求集</span><span class="sxs-lookup"><span data-stu-id="fea8c-529">API requirement sets</span></span></th>
    <th><span data-ttu-id="fea8c-530"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="fea8c-530"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-531">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="fea8c-531">Office on the web</span></span></td>
    <td><span data-ttu-id="fea8c-532">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-532">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-533">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-533">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-534">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-534">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fea8c-535">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-535">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="fea8c-536">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-536">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="fea8c-537">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-537">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-538">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-538">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fea8c-539">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-539">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-540">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-540">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-541">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-541">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="fea8c-542">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-542">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-543">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-543">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-544">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-544">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="fea8c-545">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-545">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-546">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-546">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-547">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-547">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="fea8c-548">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-548">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-549">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-549">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-550">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-550">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-551">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-551">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-552">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-552">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-553">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-553">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-554">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-554">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="fea8c-555">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-555">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-556">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="fea8c-556">Office on Windows</span></span><br><span data-ttu-id="fea8c-557">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="fea8c-557">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="fea8c-558">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-558">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-559">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-559">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-560">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-560">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fea8c-561">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-561">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="fea8c-562">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-562">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="fea8c-563">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-563">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-564">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-564">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fea8c-565">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-565">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-566">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-566">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-567">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-567">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-568">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-568">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="fea8c-569">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-569">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-570">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-570">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-571">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-571">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="fea8c-572">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-572">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-573">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-573">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-574">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-574">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="fea8c-575">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-575">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-576">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-576">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-577">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-577">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-578">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-578">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-579">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-579">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-580">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-580">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-581">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-581">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="fea8c-582">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-582">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-583">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="fea8c-583">Office 2019 on Windows</span></span><br><span data-ttu-id="fea8c-584">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-584">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-585">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-585">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-586">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-586">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-587">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-587">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fea8c-588">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-588">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="fea8c-589">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-589">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="fea8c-590">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-590">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-591">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-591">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-592">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-592">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-593">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-593">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-594">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-594">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="fea8c-595">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-595">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-596">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-596">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-597">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-597">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="fea8c-598">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-598">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-599">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-599">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-600">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-600">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="fea8c-601">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-601">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-602">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-602">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-603">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-603">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-604">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-604">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-605">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-605">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-606">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-606">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-607">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-607">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="fea8c-608">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-608">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-609">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="fea8c-609">Office 2016 on Windows</span></span><br><span data-ttu-id="fea8c-610">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-610">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-611">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-611">- TaskPane</span></span></td>
    <td><span data-ttu-id="fea8c-612">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-612">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fea8c-613">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fea8c-613">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="fea8c-614">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-614">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-615">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-615">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-616">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-616">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-617">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-617">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="fea8c-618">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-618">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-619">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-619">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-620">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-620">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="fea8c-621">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-621">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-622">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-622">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-623">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-623">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="fea8c-624">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-624">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-625">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-625">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-626">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-626">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-627">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-627">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-628">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-628">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-629">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-629">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-630">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-630">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="fea8c-631">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-631">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-632">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="fea8c-632">Office 2013 on Windows</span></span><br><span data-ttu-id="fea8c-633">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-633">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-634">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-634">- TaskPane</span></span></td>
    <td><span data-ttu-id="fea8c-635">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fea8c-635">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="fea8c-636">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-636">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-637">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-637">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-638">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-638">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-639">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-639">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="fea8c-640">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-640">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-641">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-641">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-642">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-642">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="fea8c-643">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-643">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-644">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-644">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-645">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-645">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="fea8c-646">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-646">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-647">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-647">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-648">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-648">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-649">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-649">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-650">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-650">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-651">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-651">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-652">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-652">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="fea8c-653">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-653">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-654">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="fea8c-654">Office on iPad</span></span><br><span data-ttu-id="fea8c-655">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="fea8c-655">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="fea8c-656">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-656">- TaskPane</span></span></td>
    <td><span data-ttu-id="fea8c-657">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-657">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fea8c-658">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-658">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="fea8c-659">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-659">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="fea8c-660">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-660">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-661">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-661">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-662">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-662">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-663">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-663">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-664">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-664">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="fea8c-665">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-665">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-666">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-666">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-667">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-667">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="fea8c-668">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-668">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-669">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-669">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-670">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-670">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="fea8c-671">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-671">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-672">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-672">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-673">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-673">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-674">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-674">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-675">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-675">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-676">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-676">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-677">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-677">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="fea8c-678">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-678">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-679">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="fea8c-679">Office on Mac</span></span><br><span data-ttu-id="fea8c-680">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="fea8c-680">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="fea8c-681">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-681">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-682">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-682">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-683">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-683">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fea8c-684">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-684">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="fea8c-685">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-685">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="fea8c-686">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-686">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-687">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-687">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fea8c-688">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-688">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-689">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-689">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-690">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-690">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-691">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-691">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="fea8c-692">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-692">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-693">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-693">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-694">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-694">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="fea8c-695">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-695">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-696">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-696">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-697">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-697">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="fea8c-698">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-698">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-699">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-699">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-700">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-700">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-701">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-701">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-702">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-702">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-703">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-703">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-704">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-704">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="fea8c-705">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-705">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-706">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="fea8c-706">Office 2019 on Mac</span></span><br><span data-ttu-id="fea8c-707">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-707">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-708">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-708">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-709">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-709">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-710">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-710">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fea8c-711">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-711">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="fea8c-712">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-712">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="fea8c-713">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-713">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-714">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-714">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-715">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-715">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-716">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-716">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-717">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-717">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="fea8c-718">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-718">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-719">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-719">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-720">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-720">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="fea8c-721">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-721">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-722">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-722">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-723">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-723">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="fea8c-724">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-724">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-725">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-725">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-726">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-726">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-727">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-727">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-728">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-728">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-729">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-729">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-730">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-730">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="fea8c-731">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-731">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-732">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="fea8c-732">Office 2016 on Mac</span></span><br><span data-ttu-id="fea8c-733">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-733">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-734">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-734">- TaskPane</span></span></td>
    <td><span data-ttu-id="fea8c-735">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-735">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="fea8c-736">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fea8c-736">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="fea8c-737">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-737">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-738">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-738">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="fea8c-739">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-739">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-740">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-740">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="fea8c-741">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-741">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-742">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-742">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-743">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-743">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="fea8c-744">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-744">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="fea8c-745">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-745">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="fea8c-746">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-746">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="fea8c-747">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-747">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-748">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-748">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-749">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-749">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-750">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-750">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="fea8c-751">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-751">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="fea8c-752">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-752">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="fea8c-753">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-753">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="fea8c-754">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-754">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="fea8c-755">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="fea8c-755">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="fea8c-756">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="fea8c-756">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fea8c-757">平台</span><span class="sxs-lookup"><span data-stu-id="fea8c-757">Platform</span></span></th>
    <th><span data-ttu-id="fea8c-758">扩展点</span><span class="sxs-lookup"><span data-stu-id="fea8c-758">Extension points</span></span></th>
    <th><span data-ttu-id="fea8c-759">API 要求集</span><span class="sxs-lookup"><span data-stu-id="fea8c-759">API requirement sets</span></span></th>
    <th><span data-ttu-id="fea8c-760"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="fea8c-760"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-761">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="fea8c-761">Office on the web</span></span></td>
    <td><span data-ttu-id="fea8c-762">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-762">
      - Content</span></span><br><span data-ttu-id="fea8c-763">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-763">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-764">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-764">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-765">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-765">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="fea8c-766">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-766">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-767">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-767">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fea8c-768">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-768">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-769">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-769">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="fea8c-770">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-770">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-771">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-771">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-772">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-772">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-773">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-773">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-774">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-774">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-775">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-775">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-776">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-776">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-777">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="fea8c-777">Office on Windows</span></span><br><span data-ttu-id="fea8c-778">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="fea8c-778">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="fea8c-779">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-779">
      - Content</span></span><br><span data-ttu-id="fea8c-780">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-780">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-781">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-781">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-782">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-782">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="fea8c-783">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-783">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-784">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-784">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fea8c-785">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-785">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-786">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-786">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="fea8c-787">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-787">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-788">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-788">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-789">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-789">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-790">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-790">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-791">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-791">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-792">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-792">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-793">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-793">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-794">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="fea8c-794">Office 2019 on Windows</span></span><br><span data-ttu-id="fea8c-795">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-795">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-796">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-796">
      - Content</span></span><br><span data-ttu-id="fea8c-797">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-797">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-798">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-798">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-799">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-799">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-800">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-800">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-801">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-801">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="fea8c-802">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-802">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-803">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-803">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-804">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-804">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-805">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-805">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-806">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-806">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-807">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-807">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-808">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-808">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-809">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="fea8c-809">Office 2016 on Windows</span></span><br><span data-ttu-id="fea8c-810">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-810">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-811">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-811">
      - Content</span></span><br><span data-ttu-id="fea8c-812">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-812">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="fea8c-813">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fea8c-813">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="fea8c-814">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-814">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-815">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-815">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="fea8c-816">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-816">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-817">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-817">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-818">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-818">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-819">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-819">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-820">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-820">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-821">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-821">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-822">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-822">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-823">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="fea8c-823">Office 2013 on Windows</span></span><br><span data-ttu-id="fea8c-824">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-824">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-825">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-825">
      - Content</span></span><br><span data-ttu-id="fea8c-826">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-826">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="fea8c-827">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fea8c-827">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="fea8c-828">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-828">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-829">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-829">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="fea8c-830">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-830">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-831">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-831">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-832">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-832">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-833">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-833">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-834">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-834">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-835">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-835">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-836">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-836">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-837">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="fea8c-837">Office on iPad</span></span><br><span data-ttu-id="fea8c-838">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="fea8c-838">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="fea8c-839">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-839">
      - Content</span></span><br><span data-ttu-id="fea8c-840">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-840">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="fea8c-841">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-841">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="fea8c-842">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-842">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-843">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-843">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-844">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-844">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="fea8c-845">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-845">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-846">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-846">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-847">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-847">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-848">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-848">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-849">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-849">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-850">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-850">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-851">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-851">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-852">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="fea8c-852">Office on Mac</span></span><br><span data-ttu-id="fea8c-853">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="fea8c-853">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="fea8c-854">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-854">
      - Content</span></span><br><span data-ttu-id="fea8c-855">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-855">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-856">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-856">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-857">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-857">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="fea8c-858">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-858">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-859">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-859">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="fea8c-860">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-860">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-861">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-861">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="fea8c-862">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-862">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-863">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-863">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-864">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-864">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-865">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-865">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-866">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-866">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-867">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-867">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-868">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-868">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-869">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="fea8c-869">Office 2019 on Mac</span></span><br><span data-ttu-id="fea8c-870">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-870">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-871">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-871">
      - Content</span></span><br><span data-ttu-id="fea8c-872">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-872">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-873">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-873">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-874">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-874">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-875">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-875">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-876">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-876">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="fea8c-877">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-877">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-878">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-878">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-879">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-879">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-880">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-880">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-881">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-881">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-882">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-882">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-883">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-883">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-884">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="fea8c-884">Office 2016 on Mac</span></span><br><span data-ttu-id="fea8c-885">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-885">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-886">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-886">
      - Content</span></span><br><span data-ttu-id="fea8c-887">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-887">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="fea8c-888">
       - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fea8c-888">
       - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="fea8c-889">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-889">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-890">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-890">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="fea8c-891">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-891">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="fea8c-892">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-892">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-893">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-893">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="fea8c-894">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-894">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="fea8c-895">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-895">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-896">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-896">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-897">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-897">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="fea8c-898">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="fea8c-898">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="fea8c-899">OneNote</span><span class="sxs-lookup"><span data-stu-id="fea8c-899">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fea8c-900">平台</span><span class="sxs-lookup"><span data-stu-id="fea8c-900">Platform</span></span></th>
    <th><span data-ttu-id="fea8c-901">扩展点</span><span class="sxs-lookup"><span data-stu-id="fea8c-901">Extension points</span></span></th>
    <th><span data-ttu-id="fea8c-902">API 要求集</span><span class="sxs-lookup"><span data-stu-id="fea8c-902">API requirement sets</span></span></th>
    <th><span data-ttu-id="fea8c-903"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="fea8c-903"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-904">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="fea8c-904">Office on the web</span></span></td>
    <td><span data-ttu-id="fea8c-905">
      - 内容</span><span class="sxs-lookup"><span data-stu-id="fea8c-905">
      - Content</span></span><br><span data-ttu-id="fea8c-906">
      - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-906">
      - TaskPane</span></span><br><span data-ttu-id="fea8c-907">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-907">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-908">
      - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-908">
      - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="fea8c-909">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-909">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="fea8c-910">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-910">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="fea8c-911">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-911">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="fea8c-912">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-912">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="fea8c-913">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-913">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="fea8c-914">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-914">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="fea8c-915">Project</span><span class="sxs-lookup"><span data-stu-id="fea8c-915">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fea8c-916">平台</span><span class="sxs-lookup"><span data-stu-id="fea8c-916">Platform</span></span></th>
    <th><span data-ttu-id="fea8c-917">扩展点</span><span class="sxs-lookup"><span data-stu-id="fea8c-917">Extension points</span></span></th>
    <th><span data-ttu-id="fea8c-918">API 要求集</span><span class="sxs-lookup"><span data-stu-id="fea8c-918">API requirement sets</span></span></th>
    <th><span data-ttu-id="fea8c-919"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="fea8c-919"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-920">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="fea8c-920">Office 2019 on Windows</span></span><br><span data-ttu-id="fea8c-921">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-921">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-922">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-922">- TaskPane</span></span></td>
    <td><span data-ttu-id="fea8c-923">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-923">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="fea8c-924">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-924">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-925">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-925">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-926">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="fea8c-926">Office 2016 on Windows</span></span><br><span data-ttu-id="fea8c-927">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-927">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-928">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-928">- TaskPane</span></span></td>
    <td><span data-ttu-id="fea8c-929">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-929">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="fea8c-930">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-930">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-931">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-931">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fea8c-932">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="fea8c-932">Office 2013 on Windows</span></span><br><span data-ttu-id="fea8c-933">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="fea8c-933">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="fea8c-934">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fea8c-934">- TaskPane</span></span></td>
    <td><span data-ttu-id="fea8c-935">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-935">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="fea8c-936">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="fea8c-936">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="fea8c-937">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="fea8c-937">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="fea8c-938">另请参阅</span><span class="sxs-lookup"><span data-stu-id="fea8c-938">See also</span></span>

- [<span data-ttu-id="fea8c-939">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="fea8c-939">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="fea8c-940">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="fea8c-940">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="fea8c-941">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="fea8c-941">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="fea8c-942">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="fea8c-942">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="fea8c-943">API 参考文档</span><span class="sxs-lookup"><span data-stu-id="fea8c-943">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="fea8c-944">Microsoft 365 应用版的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="fea8c-944">Update history for Microsoft 365 Apps</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="fea8c-945">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="fea8c-945">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="fea8c-946">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="fea8c-946">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="fea8c-947">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="fea8c-947">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="fea8c-948">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="fea8c-948">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="fea8c-949">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="fea8c-949">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="fea8c-950">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="fea8c-950">Developing Office Add-ins</span></span>](../develop/develop-overview.md)

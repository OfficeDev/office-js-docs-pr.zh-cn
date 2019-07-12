---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 07/11/2019
localization_priority: Priority
ms.openlocfilehash: d88f7c1b9daa201d9b6bc5cfa69ac3125bf127b1
ms.sourcegitcommit: 61f8f02193ce05da957418d938f0d94cb12c468d
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/11/2019
ms.locfileid: "35630534"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="2ddb4-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="2ddb4-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="2ddb4-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="2ddb4-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="2ddb4-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="2ddb4-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="2ddb4-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="2ddb4-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="2ddb4-108">Excel</span><span class="sxs-lookup"><span data-stu-id="2ddb4-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="2ddb4-109">平台</span><span class="sxs-lookup"><span data-stu-id="2ddb4-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="2ddb4-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="2ddb4-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="2ddb4-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="2ddb4-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="2ddb4-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="2ddb4-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="2ddb4-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-114">- TaskPane</span></span><br><span data-ttu-id="2ddb4-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-115">
        - Content</span></span><br><span data-ttu-id="2ddb4-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="2ddb4-116">
        - Custom Functions</span></span><br><span data-ttu-id="2ddb4-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="2ddb4-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="2ddb4-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2ddb4-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2ddb4-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2ddb4-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2ddb4-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2ddb4-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="2ddb4-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2ddb4-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="2ddb4-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2ddb4-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="2ddb4-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-130">
        - BindingEvents</span></span><br><span data-ttu-id="2ddb4-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-131">
        - CompressedFile</span></span><br><span data-ttu-id="2ddb4-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-132">
        - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-133">
        - File</span></span><br><span data-ttu-id="2ddb4-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-134">
        - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-136">
        - Selection</span></span><br><span data-ttu-id="2ddb4-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-137">
        - Settings</span></span><br><span data-ttu-id="2ddb4-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-138">
        - TableBindings</span></span><br><span data-ttu-id="2ddb4-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-139">
        - TableCoercion</span></span><br><span data-ttu-id="2ddb4-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-140">
        - TextBindings</span></span><br><span data-ttu-id="2ddb4-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-142">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="2ddb4-142">Office on Windows</span></span><br><span data-ttu-id="2ddb4-143">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2ddb4-144">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-144">- TaskPane</span></span><br><span data-ttu-id="2ddb4-145">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-145">
        - Content</span></span><br><span data-ttu-id="2ddb4-146">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="2ddb4-146">
        - Custom Functions</span></span><br><span data-ttu-id="2ddb4-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="2ddb4-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="2ddb4-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2ddb4-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2ddb4-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2ddb4-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2ddb4-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2ddb4-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="2ddb4-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2ddb4-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="2ddb4-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2ddb4-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="2ddb4-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-160">
        - BindingEvents</span></span><br><span data-ttu-id="2ddb4-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-161">
        - CompressedFile</span></span><br><span data-ttu-id="2ddb4-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-162">
        - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-163">
        - File</span></span><br><span data-ttu-id="2ddb4-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-164">
        - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-166">
        - Selection</span></span><br><span data-ttu-id="2ddb4-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-167">
        - Settings</span></span><br><span data-ttu-id="2ddb4-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-168">
        - TableBindings</span></span><br><span data-ttu-id="2ddb4-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-169">
        - TableCoercion</span></span><br><span data-ttu-id="2ddb4-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-170">
        - TextBindings</span></span><br><span data-ttu-id="2ddb4-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-172">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2ddb4-172">Office 2019 on Windows</span></span><br><span data-ttu-id="2ddb4-173">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="2ddb4-174">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-174">- TaskPane</span></span><br><span data-ttu-id="2ddb4-175">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-175">
        - Content</span></span><br><span data-ttu-id="2ddb4-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="2ddb4-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2ddb4-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2ddb4-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2ddb4-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2ddb4-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2ddb4-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="2ddb4-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2ddb4-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="2ddb4-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-187">- BindingEvents</span></span><br><span data-ttu-id="2ddb4-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-188">
        - CompressedFile</span></span><br><span data-ttu-id="2ddb4-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-189">
        - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-190">
        - File</span></span><br><span data-ttu-id="2ddb4-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-191">
        - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-193">
        - Selection</span></span><br><span data-ttu-id="2ddb4-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-194">
        - Settings</span></span><br><span data-ttu-id="2ddb4-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-195">
        - TableBindings</span></span><br><span data-ttu-id="2ddb4-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-196">
        - TableCoercion</span></span><br><span data-ttu-id="2ddb4-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-197">
        - TextBindings</span></span><br><span data-ttu-id="2ddb4-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-199">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2ddb4-199">Office 2016 on Windows</span></span><br><span data-ttu-id="2ddb4-200">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="2ddb4-201">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-201">- TaskPane</span></span><br><span data-ttu-id="2ddb4-202">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-202">
        - Content</span></span></td>
    <td><span data-ttu-id="2ddb4-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="2ddb4-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="2ddb4-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-206">- BindingEvents</span></span><br><span data-ttu-id="2ddb4-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-207">
        - CompressedFile</span></span><br><span data-ttu-id="2ddb4-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-208">
        - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-209">
        - File</span></span><br><span data-ttu-id="2ddb4-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-210">
        - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-212">
        - Selection</span></span><br><span data-ttu-id="2ddb4-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-213">
        - Settings</span></span><br><span data-ttu-id="2ddb4-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-214">
        - TableBindings</span></span><br><span data-ttu-id="2ddb4-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-215">
        - TableCoercion</span></span><br><span data-ttu-id="2ddb4-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-216">
        - TextBindings</span></span><br><span data-ttu-id="2ddb4-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-218">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="2ddb4-218">Office 2013 on Windows</span></span><br><span data-ttu-id="2ddb4-219">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="2ddb4-220">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-220">
        - TaskPane</span></span><br><span data-ttu-id="2ddb4-221">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="2ddb4-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="2ddb4-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="2ddb4-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-224">
        - BindingEvents</span></span><br><span data-ttu-id="2ddb4-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-225">
        - CompressedFile</span></span><br><span data-ttu-id="2ddb4-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-226">
        - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-227">
        - File</span></span><br><span data-ttu-id="2ddb4-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-228">
        - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-230">
        - Selection</span></span><br><span data-ttu-id="2ddb4-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-231">
        - Settings</span></span><br><span data-ttu-id="2ddb4-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-232">
        - TableBindings</span></span><br><span data-ttu-id="2ddb4-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-233">
        - TableCoercion</span></span><br><span data-ttu-id="2ddb4-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-234">
        - TextBindings</span></span><br><span data-ttu-id="2ddb4-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-236">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="2ddb4-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="2ddb4-237">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="2ddb4-238">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-238">- TaskPane</span></span><br><span data-ttu-id="2ddb4-239">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-239">
        - Content</span></span><br><span data-ttu-id="2ddb4-240">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="2ddb4-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="2ddb4-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2ddb4-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2ddb4-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2ddb4-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2ddb4-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2ddb4-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="2ddb4-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2ddb4-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="2ddb4-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="2ddb4-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-252">- BindingEvents</span></span><br><span data-ttu-id="2ddb4-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-253">
        - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-254">
        - File</span></span><br><span data-ttu-id="2ddb4-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-255">
        - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-257">
        - Selection</span></span><br><span data-ttu-id="2ddb4-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-258">
        - Settings</span></span><br><span data-ttu-id="2ddb4-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-259">
        - TableBindings</span></span><br><span data-ttu-id="2ddb4-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-260">
        - TableCoercion</span></span><br><span data-ttu-id="2ddb4-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-261">
        - TextBindings</span></span><br><span data-ttu-id="2ddb4-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-263">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="2ddb4-263">Office apps on Mac</span></span><br><span data-ttu-id="2ddb4-264">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="2ddb4-265">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-265">- TaskPane</span></span><br><span data-ttu-id="2ddb4-266">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-266">
        - Content</span></span><br><span data-ttu-id="2ddb4-267">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="2ddb4-267">
        - Custom Functions</span></span><br><span data-ttu-id="2ddb4-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="2ddb4-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2ddb4-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2ddb4-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2ddb4-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2ddb4-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2ddb4-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="2ddb4-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2ddb4-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="2ddb4-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2ddb4-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="2ddb4-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-281">- BindingEvents</span></span><br><span data-ttu-id="2ddb4-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-282">
        - CompressedFile</span></span><br><span data-ttu-id="2ddb4-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-283">
        - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-284">
        - File</span></span><br><span data-ttu-id="2ddb4-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-285">
        - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-287">
        - PdfFile</span></span><br><span data-ttu-id="2ddb4-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-288">
        - Selection</span></span><br><span data-ttu-id="2ddb4-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-289">
        - Settings</span></span><br><span data-ttu-id="2ddb4-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-290">
        - TableBindings</span></span><br><span data-ttu-id="2ddb4-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-291">
        - TableCoercion</span></span><br><span data-ttu-id="2ddb4-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-292">
        - TextBindings</span></span><br><span data-ttu-id="2ddb4-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-294">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2ddb4-294">Office 2019 for Mac</span></span><br><span data-ttu-id="2ddb4-295">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="2ddb4-296">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-296">- TaskPane</span></span><br><span data-ttu-id="2ddb4-297">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-297">
        - Content</span></span><br><span data-ttu-id="2ddb4-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="2ddb4-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2ddb4-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2ddb4-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2ddb4-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2ddb4-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2ddb4-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="2ddb4-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2ddb4-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="2ddb4-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-309">- BindingEvents</span></span><br><span data-ttu-id="2ddb4-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-310">
        - CompressedFile</span></span><br><span data-ttu-id="2ddb4-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-311">
        - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-312">
        - File</span></span><br><span data-ttu-id="2ddb4-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-313">
        - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-315">
        - PdfFile</span></span><br><span data-ttu-id="2ddb4-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-316">
        - Selection</span></span><br><span data-ttu-id="2ddb4-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-317">
        - Settings</span></span><br><span data-ttu-id="2ddb4-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-318">
        - TableBindings</span></span><br><span data-ttu-id="2ddb4-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-319">
        - TableCoercion</span></span><br><span data-ttu-id="2ddb4-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-320">
        - TextBindings</span></span><br><span data-ttu-id="2ddb4-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-322">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2ddb4-322">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="2ddb4-323">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="2ddb4-324">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-324">- TaskPane</span></span><br><span data-ttu-id="2ddb4-325">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-325">
        - Content</span></span></td>
    <td><span data-ttu-id="2ddb4-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="2ddb4-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="2ddb4-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-329">- BindingEvents</span></span><br><span data-ttu-id="2ddb4-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-330">
        - CompressedFile</span></span><br><span data-ttu-id="2ddb4-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-331">
        - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-332">
        - File</span></span><br><span data-ttu-id="2ddb4-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-333">
        - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-335">
        - PdfFile</span></span><br><span data-ttu-id="2ddb4-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-336">
        - Selection</span></span><br><span data-ttu-id="2ddb4-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-337">
        - Settings</span></span><br><span data-ttu-id="2ddb4-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-338">
        - TableBindings</span></span><br><span data-ttu-id="2ddb4-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-339">
        - TableCoercion</span></span><br><span data-ttu-id="2ddb4-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-340">
        - TextBindings</span></span><br><span data-ttu-id="2ddb4-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="2ddb4-342">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="2ddb4-343">自定义函数</span><span class="sxs-lookup"><span data-stu-id="2ddb4-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="2ddb4-344">平台</span><span class="sxs-lookup"><span data-stu-id="2ddb4-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="2ddb4-345">扩展点</span><span class="sxs-lookup"><span data-stu-id="2ddb4-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="2ddb4-346">API 要求集</span><span class="sxs-lookup"><span data-stu-id="2ddb4-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="2ddb4-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-348">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="2ddb4-348">Office on the web</span></span></td>
    <td><span data-ttu-id="2ddb4-349">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="2ddb4-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="2ddb4-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-351">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="2ddb4-351">Office on Windows</span></span><br><span data-ttu-id="2ddb4-352">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="2ddb4-353">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="2ddb4-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="2ddb4-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-355">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="2ddb4-355">Office for Mac</span></span><br><span data-ttu-id="2ddb4-356">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="2ddb4-357">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="2ddb4-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="2ddb4-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="2ddb4-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="2ddb4-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="2ddb4-360">平台</span><span class="sxs-lookup"><span data-stu-id="2ddb4-360">Platform</span></span></th>
    <th><span data-ttu-id="2ddb4-361">扩展点</span><span class="sxs-lookup"><span data-stu-id="2ddb4-361">Extension points</span></span></th>
    <th><span data-ttu-id="2ddb4-362">API 要求集</span><span class="sxs-lookup"><span data-stu-id="2ddb4-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="2ddb4-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-364">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="2ddb4-364">Office on the web</span></span><br><span data-ttu-id="2ddb4-365">（新）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-365">New</span></span></td>
    <td> <span data-ttu-id="2ddb4-366">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2ddb4-366">- Mail Read</span></span><br><span data-ttu-id="2ddb4-367">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2ddb4-367">
      - Mail Compose</span></span><br><span data-ttu-id="2ddb4-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2ddb4-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2ddb4-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2ddb4-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2ddb4-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2ddb4-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="2ddb4-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="2ddb4-376">不可用</span><span class="sxs-lookup"><span data-stu-id="2ddb4-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-377">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="2ddb4-377">Office on the web</span></span><br><span data-ttu-id="2ddb4-378">（经典）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-378">Classic.</span></span></td>
    <td> <span data-ttu-id="2ddb4-379">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2ddb4-379">- Mail Read</span></span><br><span data-ttu-id="2ddb4-380">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2ddb4-380">
      - Mail Compose</span></span><br><span data-ttu-id="2ddb4-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2ddb4-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2ddb4-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2ddb4-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2ddb4-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2ddb4-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="2ddb4-388">不可用</span><span class="sxs-lookup"><span data-stu-id="2ddb4-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-389">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="2ddb4-389">Office on Windows</span></span><br><span data-ttu-id="2ddb4-390">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2ddb4-391">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2ddb4-391">- Mail Read</span></span><br><span data-ttu-id="2ddb4-392">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2ddb4-392">
      - Mail Compose</span></span><br><span data-ttu-id="2ddb4-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="2ddb4-394">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="2ddb4-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="2ddb4-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2ddb4-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2ddb4-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2ddb4-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2ddb4-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2ddb4-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="2ddb4-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="2ddb4-402">不可用</span><span class="sxs-lookup"><span data-stu-id="2ddb4-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-403">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2ddb4-403">Office 2019 on Windows</span></span><br><span data-ttu-id="2ddb4-404">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-405">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2ddb4-405">- Mail Read</span></span><br><span data-ttu-id="2ddb4-406">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2ddb4-406">
      - Mail Compose</span></span><br><span data-ttu-id="2ddb4-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="2ddb4-408">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="2ddb4-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="2ddb4-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2ddb4-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2ddb4-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2ddb4-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2ddb4-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2ddb4-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="2ddb4-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="2ddb4-416">不可用</span><span class="sxs-lookup"><span data-stu-id="2ddb4-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-417">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2ddb4-417">Office 2016 on Windows</span></span><br><span data-ttu-id="2ddb4-418">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-419">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2ddb4-419">- Mail Read</span></span><br><span data-ttu-id="2ddb4-420">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2ddb4-420">
      - Mail Compose</span></span><br><span data-ttu-id="2ddb4-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="2ddb4-422">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="2ddb4-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="2ddb4-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2ddb4-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2ddb4-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2ddb4-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="2ddb4-427">不可用</span><span class="sxs-lookup"><span data-stu-id="2ddb4-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-428">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="2ddb4-428">Office 2013 on Windows</span></span><br><span data-ttu-id="2ddb4-429">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-430">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2ddb4-430">- Mail Read</span></span><br><span data-ttu-id="2ddb4-431">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2ddb4-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="2ddb4-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2ddb4-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2ddb4-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="2ddb4-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="2ddb4-436">不可用</span><span class="sxs-lookup"><span data-stu-id="2ddb4-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-437">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="2ddb4-437">Office apps on iOS</span></span><br><span data-ttu-id="2ddb4-438">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2ddb4-439">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2ddb4-439">- Mail Read</span></span><br><span data-ttu-id="2ddb4-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2ddb4-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2ddb4-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2ddb4-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2ddb4-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="2ddb4-446">不可用</span><span class="sxs-lookup"><span data-stu-id="2ddb4-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-447">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="2ddb4-447">Office apps on Mac</span></span><br><span data-ttu-id="2ddb4-448">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2ddb4-449">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2ddb4-449">- Mail Read</span></span><br><span data-ttu-id="2ddb4-450">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2ddb4-450">
      - Mail Compose</span></span><br><span data-ttu-id="2ddb4-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2ddb4-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2ddb4-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2ddb4-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2ddb4-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2ddb4-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="2ddb4-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="2ddb4-459">不可用</span><span class="sxs-lookup"><span data-stu-id="2ddb4-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-460">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2ddb4-460">Office 2019 for Mac</span></span><br><span data-ttu-id="2ddb4-461">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-462">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2ddb4-462">- Mail Read</span></span><br><span data-ttu-id="2ddb4-463">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2ddb4-463">
      - Mail Compose</span></span><br><span data-ttu-id="2ddb4-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2ddb4-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2ddb4-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2ddb4-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2ddb4-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2ddb4-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="2ddb4-471">不可用</span><span class="sxs-lookup"><span data-stu-id="2ddb4-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-472">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2ddb4-472">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="2ddb4-473">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-474">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2ddb4-474">- Mail Read</span></span><br><span data-ttu-id="2ddb4-475">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="2ddb4-475">
      - Mail Compose</span></span><br><span data-ttu-id="2ddb4-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2ddb4-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2ddb4-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2ddb4-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2ddb4-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2ddb4-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="2ddb4-483">不可用</span><span class="sxs-lookup"><span data-stu-id="2ddb4-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-484">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="2ddb4-484">Office apps on Android</span></span><br><span data-ttu-id="2ddb4-485">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2ddb4-486">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="2ddb4-486">- Mail Read</span></span><br><span data-ttu-id="2ddb4-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2ddb4-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2ddb4-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2ddb4-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2ddb4-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="2ddb4-493">不可用</span><span class="sxs-lookup"><span data-stu-id="2ddb4-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="2ddb4-494">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="2ddb4-495">Word</span><span class="sxs-lookup"><span data-stu-id="2ddb4-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="2ddb4-496">平台</span><span class="sxs-lookup"><span data-stu-id="2ddb4-496">Platform</span></span></th>
    <th><span data-ttu-id="2ddb4-497">扩展点</span><span class="sxs-lookup"><span data-stu-id="2ddb4-497">Extension points</span></span></th>
    <th><span data-ttu-id="2ddb4-498">API 要求集</span><span class="sxs-lookup"><span data-stu-id="2ddb4-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="2ddb4-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-500">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="2ddb4-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="2ddb4-501">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-501">- TaskPane</span></span><br><span data-ttu-id="2ddb4-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2ddb4-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2ddb4-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2ddb4-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-509">- BindingEvents</span></span><br><span data-ttu-id="2ddb4-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2ddb4-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="2ddb4-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-511">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-512">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-512">
         - File</span></span><br><span data-ttu-id="2ddb4-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="2ddb4-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-514">
         - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2ddb4-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-517">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-518">
         - Selection</span></span><br><span data-ttu-id="2ddb4-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-519">
         - Settings</span></span><br><span data-ttu-id="2ddb4-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-520">
         - TableBindings</span></span><br><span data-ttu-id="2ddb4-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-521">
         - TableCoercion</span></span><br><span data-ttu-id="2ddb4-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-522">
         - TextBindings</span></span><br><span data-ttu-id="2ddb4-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-523">
         - TextCoercion</span></span><br><span data-ttu-id="2ddb4-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-525">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="2ddb4-525">Office on Windows</span></span><br><span data-ttu-id="2ddb4-526">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2ddb4-527">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-527">- TaskPane</span></span><br><span data-ttu-id="2ddb4-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2ddb4-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2ddb4-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2ddb4-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-535">- BindingEvents</span></span><br><span data-ttu-id="2ddb4-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-536">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2ddb4-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="2ddb4-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-538">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-539">
         - File</span></span><br><span data-ttu-id="2ddb4-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="2ddb4-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-541">
         - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2ddb4-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-544">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-545">
         - Selection</span></span><br><span data-ttu-id="2ddb4-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-546">
         - Settings</span></span><br><span data-ttu-id="2ddb4-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-547">
         - TableBindings</span></span><br><span data-ttu-id="2ddb4-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-548">
         - TableCoercion</span></span><br><span data-ttu-id="2ddb4-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-549">
         - TextBindings</span></span><br><span data-ttu-id="2ddb4-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-550">
         - TextCoercion</span></span><br><span data-ttu-id="2ddb4-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-552">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2ddb4-552">Office 2019 on Windows</span></span><br><span data-ttu-id="2ddb4-553">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-554">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-554">- TaskPane</span></span><br><span data-ttu-id="2ddb4-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2ddb4-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2ddb4-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-561">- BindingEvents</span></span><br><span data-ttu-id="2ddb4-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-562">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2ddb4-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="2ddb4-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-564">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-565">
         - File</span></span><br><span data-ttu-id="2ddb4-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="2ddb4-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-567">
         - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2ddb4-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-570">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-571">
         - Selection</span></span><br><span data-ttu-id="2ddb4-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-572">
         - Settings</span></span><br><span data-ttu-id="2ddb4-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-573">
         - TableBindings</span></span><br><span data-ttu-id="2ddb4-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-574">
         - TableCoercion</span></span><br><span data-ttu-id="2ddb4-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-575">
         - TextBindings</span></span><br><span data-ttu-id="2ddb4-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-576">
         - TextCoercion</span></span><br><span data-ttu-id="2ddb4-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-578">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2ddb4-578">Office 2016 on Windows</span></span><br><span data-ttu-id="2ddb4-579">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-580">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="2ddb4-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="2ddb4-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-584">- BindingEvents</span></span><br><span data-ttu-id="2ddb4-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-585">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2ddb4-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="2ddb4-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-587">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-588">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-588">
         - File</span></span><br><span data-ttu-id="2ddb4-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="2ddb4-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-590">
         - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2ddb4-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-593">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-594">
         - Selection</span></span><br><span data-ttu-id="2ddb4-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-595">
         - Settings</span></span><br><span data-ttu-id="2ddb4-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-596">
         - TableBindings</span></span><br><span data-ttu-id="2ddb4-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-597">
         - TableCoercion</span></span><br><span data-ttu-id="2ddb4-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-598">
         - TextBindings</span></span><br><span data-ttu-id="2ddb4-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-599">
         - TextCoercion</span></span><br><span data-ttu-id="2ddb4-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-601">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="2ddb4-601">Office 2013 on Windows</span></span><br><span data-ttu-id="2ddb4-602">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-603">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="2ddb4-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="2ddb4-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-606">- BindingEvents</span></span><br><span data-ttu-id="2ddb4-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-607">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2ddb4-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="2ddb4-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-609">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-610">
         - File</span></span><br><span data-ttu-id="2ddb4-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="2ddb4-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-612">
         - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2ddb4-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-615">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-616">
         - Selection</span></span><br><span data-ttu-id="2ddb4-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-617">
         - Settings</span></span><br><span data-ttu-id="2ddb4-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-618">
         - TableBindings</span></span><br><span data-ttu-id="2ddb4-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-619">
         - TableCoercion</span></span><br><span data-ttu-id="2ddb4-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-620">
         - TextBindings</span></span><br><span data-ttu-id="2ddb4-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-621">
         - TextCoercion</span></span><br><span data-ttu-id="2ddb4-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-623">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="2ddb4-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="2ddb4-624">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2ddb4-625">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="2ddb4-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2ddb4-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2ddb4-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="2ddb4-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-631">- BindingEvents</span></span><br><span data-ttu-id="2ddb4-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-632">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2ddb4-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="2ddb4-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-634">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-635">
         - File</span></span><br><span data-ttu-id="2ddb4-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="2ddb4-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-637">
         - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2ddb4-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-640">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-641">
         - Selection</span></span><br><span data-ttu-id="2ddb4-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-642">
         - Settings</span></span><br><span data-ttu-id="2ddb4-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-643">
         - TableBindings</span></span><br><span data-ttu-id="2ddb4-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-644">
         - TableCoercion</span></span><br><span data-ttu-id="2ddb4-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-645">
         - TextBindings</span></span><br><span data-ttu-id="2ddb4-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-646">
         - TextCoercion</span></span><br><span data-ttu-id="2ddb4-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-648">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="2ddb4-648">Office apps on Mac</span></span><br><span data-ttu-id="2ddb4-649">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2ddb4-650">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-650">- TaskPane</span></span><br><span data-ttu-id="2ddb4-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2ddb4-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2ddb4-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2ddb4-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="2ddb4-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-658">- BindingEvents</span></span><br><span data-ttu-id="2ddb4-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-659">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2ddb4-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="2ddb4-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-661">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-662">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-662">
         - File</span></span><br><span data-ttu-id="2ddb4-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="2ddb4-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-664">
         - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2ddb4-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-667">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-668">
         - Selection</span></span><br><span data-ttu-id="2ddb4-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-669">
         - Settings</span></span><br><span data-ttu-id="2ddb4-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-670">
         - TableBindings</span></span><br><span data-ttu-id="2ddb4-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-671">
         - TableCoercion</span></span><br><span data-ttu-id="2ddb4-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-672">
         - TextBindings</span></span><br><span data-ttu-id="2ddb4-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-673">
         - TextCoercion</span></span><br><span data-ttu-id="2ddb4-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-675">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2ddb4-675">Office 2019 for Mac</span></span><br><span data-ttu-id="2ddb4-676">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-677">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-677">- TaskPane</span></span><br><span data-ttu-id="2ddb4-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2ddb4-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2ddb4-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="2ddb4-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-684">- BindingEvents</span></span><br><span data-ttu-id="2ddb4-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-685">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2ddb4-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="2ddb4-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-687">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-688">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-688">
         - File</span></span><br><span data-ttu-id="2ddb4-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="2ddb4-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-690">
         - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2ddb4-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-693">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-694">
         - Selection</span></span><br><span data-ttu-id="2ddb4-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-695">
         - Settings</span></span><br><span data-ttu-id="2ddb4-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-696">
         - TableBindings</span></span><br><span data-ttu-id="2ddb4-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-697">
         - TableCoercion</span></span><br><span data-ttu-id="2ddb4-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-698">
         - TextBindings</span></span><br><span data-ttu-id="2ddb4-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-699">
         - TextCoercion</span></span><br><span data-ttu-id="2ddb4-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-701">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2ddb4-701">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="2ddb4-702">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-703">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="2ddb4-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="2ddb4-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-707">- BindingEvents</span></span><br><span data-ttu-id="2ddb4-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-708">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2ddb4-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="2ddb4-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-710">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-711">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-711">
         - File</span></span><br><span data-ttu-id="2ddb4-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="2ddb4-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-713">
         - MatrixBindings</span></span><br><span data-ttu-id="2ddb4-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="2ddb4-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="2ddb4-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-716">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-717">
         - Selection</span></span><br><span data-ttu-id="2ddb4-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-718">
         - Settings</span></span><br><span data-ttu-id="2ddb4-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-719">
         - TableBindings</span></span><br><span data-ttu-id="2ddb4-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-720">
         - TableCoercion</span></span><br><span data-ttu-id="2ddb4-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-721">
         - TextBindings</span></span><br><span data-ttu-id="2ddb4-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-722">
         - TextCoercion</span></span><br><span data-ttu-id="2ddb4-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="2ddb4-724">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="2ddb4-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="2ddb4-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="2ddb4-726">平台</span><span class="sxs-lookup"><span data-stu-id="2ddb4-726">Platform</span></span></th>
    <th><span data-ttu-id="2ddb4-727">扩展点</span><span class="sxs-lookup"><span data-stu-id="2ddb4-727">Extension points</span></span></th>
    <th><span data-ttu-id="2ddb4-728">API 要求集</span><span class="sxs-lookup"><span data-stu-id="2ddb4-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="2ddb4-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-730">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="2ddb4-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="2ddb4-731">- 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-731">- Content</span></span><br><span data-ttu-id="2ddb4-732">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-732">
         - TaskPane</span></span><br><span data-ttu-id="2ddb4-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2ddb4-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2ddb4-737">- ActiveView</span></span><br><span data-ttu-id="2ddb4-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-738">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-739">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-740">
         - File</span></span><br><span data-ttu-id="2ddb4-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-741">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-742">
         - Selection</span></span><br><span data-ttu-id="2ddb4-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-743">
         - Settings</span></span><br><span data-ttu-id="2ddb4-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-745">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="2ddb4-745">Office on Windows</span></span><br><span data-ttu-id="2ddb4-746">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2ddb4-747">- 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-747">- Content</span></span><br><span data-ttu-id="2ddb4-748">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-748">
         - TaskPane</span></span><br><span data-ttu-id="2ddb4-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2ddb4-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2ddb4-753">- ActiveView</span></span><br><span data-ttu-id="2ddb4-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-754">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-755">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-756">
         - File</span></span><br><span data-ttu-id="2ddb4-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-757">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-758">
         - Selection</span></span><br><span data-ttu-id="2ddb4-759">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-759">
         - Settings</span></span><br><span data-ttu-id="2ddb4-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-761">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2ddb4-761">Office 2019 on Windows</span></span><br><span data-ttu-id="2ddb4-762">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-763">- 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-763">- Content</span></span><br><span data-ttu-id="2ddb4-764">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-764">
         - TaskPane</span></span><br><span data-ttu-id="2ddb4-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2ddb4-768">- ActiveView</span></span><br><span data-ttu-id="2ddb4-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-769">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-770">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-771">
         - File</span></span><br><span data-ttu-id="2ddb4-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-772">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-773">
         - Selection</span></span><br><span data-ttu-id="2ddb4-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-774">
         - Settings</span></span><br><span data-ttu-id="2ddb4-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-776">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2ddb4-776">Office 2016 on Windows</span></span><br><span data-ttu-id="2ddb4-777">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-778">- 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-778">- Content</span></span><br><span data-ttu-id="2ddb4-779">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="2ddb4-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="2ddb4-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2ddb4-782">- ActiveView</span></span><br><span data-ttu-id="2ddb4-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-783">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-784">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-785">
         - File</span></span><br><span data-ttu-id="2ddb4-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-786">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-787">
         - Selection</span></span><br><span data-ttu-id="2ddb4-788">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-788">
         - Settings</span></span><br><span data-ttu-id="2ddb4-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-790">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="2ddb4-790">Office 2013 on Windows</span></span><br><span data-ttu-id="2ddb4-791">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-792">- 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-792">- Content</span></span><br><span data-ttu-id="2ddb4-793">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="2ddb4-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="2ddb4-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2ddb4-796">- ActiveView</span></span><br><span data-ttu-id="2ddb4-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-797">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-798">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-799">
         - File</span></span><br><span data-ttu-id="2ddb4-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-800">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-801">
         - Selection</span></span><br><span data-ttu-id="2ddb4-802">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-802">
         - Settings</span></span><br><span data-ttu-id="2ddb4-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-804">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="2ddb4-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="2ddb4-805">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2ddb4-806">- 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-806">- Content</span></span><br><span data-ttu-id="2ddb4-807">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="2ddb4-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2ddb4-810">- ActiveView</span></span><br><span data-ttu-id="2ddb4-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-811">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-812">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-813">
         - File</span></span><br><span data-ttu-id="2ddb4-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-814">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-815">
         - Selection</span></span><br><span data-ttu-id="2ddb4-816">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-816">
         - Settings</span></span><br><span data-ttu-id="2ddb4-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-818">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="2ddb4-818">Office apps on Mac</span></span><br><span data-ttu-id="2ddb4-819">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="2ddb4-820">- 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-820">- Content</span></span><br><span data-ttu-id="2ddb4-821">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-821">
         - TaskPane</span></span><br><span data-ttu-id="2ddb4-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="2ddb4-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-826">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2ddb4-826">- ActiveView</span></span><br><span data-ttu-id="2ddb4-827">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-827">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-828">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-828">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-829">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-829">
         - File</span></span><br><span data-ttu-id="2ddb4-830">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-830">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-831">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-831">
         - Selection</span></span><br><span data-ttu-id="2ddb4-832">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-832">
         - Settings</span></span><br><span data-ttu-id="2ddb4-833">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-833">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-834">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2ddb4-834">Office 2019 for Mac</span></span><br><span data-ttu-id="2ddb4-835">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-835">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-836">- 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-836">- Content</span></span><br><span data-ttu-id="2ddb4-837">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-837">
         - TaskPane</span></span><br><span data-ttu-id="2ddb4-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-841">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2ddb4-841">- ActiveView</span></span><br><span data-ttu-id="2ddb4-842">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-842">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-843">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-843">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-844">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-844">
         - File</span></span><br><span data-ttu-id="2ddb4-845">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-845">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-846">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-846">
         - Selection</span></span><br><span data-ttu-id="2ddb4-847">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-847">
         - Settings</span></span><br><span data-ttu-id="2ddb4-848">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-848">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-849">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2ddb4-849">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="2ddb4-850">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-850">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-851">- 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-851">- Content</span></span><br><span data-ttu-id="2ddb4-852">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-852">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="2ddb4-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="2ddb4-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2ddb4-855">- ActiveView</span></span><br><span data-ttu-id="2ddb4-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-856">
         - CompressedFile</span></span><br><span data-ttu-id="2ddb4-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-857">
         - DocumentEvents</span></span><br><span data-ttu-id="2ddb4-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="2ddb4-858">
         - File</span></span><br><span data-ttu-id="2ddb4-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2ddb4-859">
         - PdfFile</span></span><br><span data-ttu-id="2ddb4-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-860">
         - Selection</span></span><br><span data-ttu-id="2ddb4-861">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-861">
         - Settings</span></span><br><span data-ttu-id="2ddb4-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-862">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="2ddb4-863">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="2ddb4-863">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="2ddb4-864">OneNote</span><span class="sxs-lookup"><span data-stu-id="2ddb4-864">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="2ddb4-865">平台</span><span class="sxs-lookup"><span data-stu-id="2ddb4-865">Platform</span></span></th>
    <th><span data-ttu-id="2ddb4-866">扩展点</span><span class="sxs-lookup"><span data-stu-id="2ddb4-866">Extension points</span></span></th>
    <th><span data-ttu-id="2ddb4-867">API 要求集</span><span class="sxs-lookup"><span data-stu-id="2ddb4-867">API requirement sets</span></span></th>
    <th><span data-ttu-id="2ddb4-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-869">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="2ddb4-869">Office on the web</span></span></td>
    <td> <span data-ttu-id="2ddb4-870">- 内容</span><span class="sxs-lookup"><span data-stu-id="2ddb4-870">- Content</span></span><br><span data-ttu-id="2ddb4-871">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-871">
         - TaskPane</span></span><br><span data-ttu-id="2ddb4-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="2ddb4-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-876">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2ddb4-876">- DocumentEvents</span></span><br><span data-ttu-id="2ddb4-877">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-877">
         - HtmlCoercion</span></span><br><span data-ttu-id="2ddb4-878">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="2ddb4-878">
         - Settings</span></span><br><span data-ttu-id="2ddb4-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-879">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="2ddb4-880">项目</span><span class="sxs-lookup"><span data-stu-id="2ddb4-880">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="2ddb4-881">平台</span><span class="sxs-lookup"><span data-stu-id="2ddb4-881">Platform</span></span></th>
    <th><span data-ttu-id="2ddb4-882">扩展点</span><span class="sxs-lookup"><span data-stu-id="2ddb4-882">Extension points</span></span></th>
    <th><span data-ttu-id="2ddb4-883">API 要求集</span><span class="sxs-lookup"><span data-stu-id="2ddb4-883">API requirement sets</span></span></th>
    <th><span data-ttu-id="2ddb4-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-885">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2ddb4-885">Office 2019 on Windows</span></span><br><span data-ttu-id="2ddb4-886">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-886">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-887">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-887">- TaskPane</span></span></td>
    <td> <span data-ttu-id="2ddb4-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-889">- Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-889">- Selection</span></span><br><span data-ttu-id="2ddb4-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-890">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-891">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2ddb4-891">Office 2016 on Windows</span></span><br><span data-ttu-id="2ddb4-892">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-893">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="2ddb4-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-895">- Selection</span></span><br><span data-ttu-id="2ddb4-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2ddb4-897">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="2ddb4-897">Office 2013 on Windows</span></span><br><span data-ttu-id="2ddb4-898">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="2ddb4-899">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="2ddb4-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="2ddb4-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2ddb4-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2ddb4-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="2ddb4-901">- Selection</span></span><br><span data-ttu-id="2ddb4-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2ddb4-902">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="2ddb4-903">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2ddb4-903">See also</span></span>

- [<span data-ttu-id="2ddb4-904">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="2ddb4-904">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="2ddb4-905">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="2ddb4-905">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="2ddb4-906">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="2ddb4-906">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="2ddb4-907">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="2ddb4-907">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="2ddb4-908">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="2ddb4-908">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="2ddb4-909">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="2ddb4-909">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="2ddb4-910">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-910">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="2ddb4-911">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="2ddb4-911">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="2ddb4-912">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="2ddb4-912">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="2ddb4-913">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="2ddb4-913">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="2ddb4-914">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="2ddb4-914">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)

---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 10/09/2019
localization_priority: Priority
ms.openlocfilehash: 28d63866a03bcae99829d3a6b6c6198059a92bdc
ms.sourcegitcommit: 4d9f3e177b0bcd62804d5045f52b03e441af244f
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/10/2019
ms.locfileid: "37440148"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="3b6ba-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="3b6ba-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="3b6ba-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="3b6ba-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="3b6ba-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="3b6ba-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="3b6ba-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="3b6ba-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="3b6ba-108">Excel</span><span class="sxs-lookup"><span data-stu-id="3b6ba-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="3b6ba-109">平台</span><span class="sxs-lookup"><span data-stu-id="3b6ba-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="3b6ba-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="3b6ba-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="3b6ba-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="3b6ba-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="3b6ba-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="3b6ba-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="3b6ba-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-114">- TaskPane</span></span><br><span data-ttu-id="3b6ba-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-115">
        - Content</span></span><br><span data-ttu-id="3b6ba-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="3b6ba-116">
        - Custom Functions</span></span><br><span data-ttu-id="3b6ba-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="3b6ba-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3b6ba-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3b6ba-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3b6ba-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3b6ba-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3b6ba-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3b6ba-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3b6ba-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3b6ba-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3b6ba-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="3b6ba-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-128">
        - BindingEvents</span></span><br><span data-ttu-id="3b6ba-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-129">
        - CompressedFile</span></span><br><span data-ttu-id="3b6ba-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-130">
        - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-131">
        - File</span></span><br><span data-ttu-id="3b6ba-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-132">
        - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-134">
        - Selection</span></span><br><span data-ttu-id="3b6ba-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-135">
        - Settings</span></span><br><span data-ttu-id="3b6ba-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-136">
        - TableBindings</span></span><br><span data-ttu-id="3b6ba-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-137">
        - TableCoercion</span></span><br><span data-ttu-id="3b6ba-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-138">
        - TextBindings</span></span><br><span data-ttu-id="3b6ba-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-140">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="3b6ba-140">Office on Windows</span></span><br><span data-ttu-id="3b6ba-141">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-141">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3b6ba-142">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-142">- TaskPane</span></span><br><span data-ttu-id="3b6ba-143">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-143">
        - Content</span></span><br><span data-ttu-id="3b6ba-144">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="3b6ba-144">
        - Custom Functions</span></span><br><span data-ttu-id="3b6ba-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="3b6ba-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3b6ba-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3b6ba-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3b6ba-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3b6ba-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3b6ba-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3b6ba-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3b6ba-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3b6ba-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3b6ba-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3b6ba-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="3b6ba-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-158">
        - BindingEvents</span></span><br><span data-ttu-id="3b6ba-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-159">
        - CompressedFile</span></span><br><span data-ttu-id="3b6ba-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-160">
        - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-161">
        - File</span></span><br><span data-ttu-id="3b6ba-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-162">
        - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-164">
        - Selection</span></span><br><span data-ttu-id="3b6ba-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-165">
        - Settings</span></span><br><span data-ttu-id="3b6ba-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-166">
        - TableBindings</span></span><br><span data-ttu-id="3b6ba-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-167">
        - TableCoercion</span></span><br><span data-ttu-id="3b6ba-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-168">
        - TextBindings</span></span><br><span data-ttu-id="3b6ba-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-170">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="3b6ba-170">Office 2019 on Windows</span></span><br><span data-ttu-id="3b6ba-171">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3b6ba-172">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-172">- TaskPane</span></span><br><span data-ttu-id="3b6ba-173">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-173">
        - Content</span></span><br><span data-ttu-id="3b6ba-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="3b6ba-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3b6ba-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3b6ba-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3b6ba-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3b6ba-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3b6ba-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3b6ba-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3b6ba-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3b6ba-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-185">- BindingEvents</span></span><br><span data-ttu-id="3b6ba-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-186">
        - CompressedFile</span></span><br><span data-ttu-id="3b6ba-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-187">
        - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-188">
        - File</span></span><br><span data-ttu-id="3b6ba-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-189">
        - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-191">
        - Selection</span></span><br><span data-ttu-id="3b6ba-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-192">
        - Settings</span></span><br><span data-ttu-id="3b6ba-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-193">
        - TableBindings</span></span><br><span data-ttu-id="3b6ba-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-194">
        - TableCoercion</span></span><br><span data-ttu-id="3b6ba-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-195">
        - TextBindings</span></span><br><span data-ttu-id="3b6ba-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-197">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="3b6ba-197">Office 2016 on Windows</span></span><br><span data-ttu-id="3b6ba-198">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3b6ba-199">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-199">- TaskPane</span></span><br><span data-ttu-id="3b6ba-200">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-200">
        - Content</span></span></td>
    <td><span data-ttu-id="3b6ba-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3b6ba-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3b6ba-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-204">- BindingEvents</span></span><br><span data-ttu-id="3b6ba-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-205">
        - CompressedFile</span></span><br><span data-ttu-id="3b6ba-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-206">
        - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-207">
        - File</span></span><br><span data-ttu-id="3b6ba-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-208">
        - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-210">
        - Selection</span></span><br><span data-ttu-id="3b6ba-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-211">
        - Settings</span></span><br><span data-ttu-id="3b6ba-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-212">
        - TableBindings</span></span><br><span data-ttu-id="3b6ba-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-213">
        - TableCoercion</span></span><br><span data-ttu-id="3b6ba-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-214">
        - TextBindings</span></span><br><span data-ttu-id="3b6ba-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-216">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="3b6ba-216">Office 2013 on Windows</span></span><br><span data-ttu-id="3b6ba-217">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3b6ba-218">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-218">
        - TaskPane</span></span><br><span data-ttu-id="3b6ba-219">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="3b6ba-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="3b6ba-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3b6ba-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-222">
        - BindingEvents</span></span><br><span data-ttu-id="3b6ba-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-223">
        - CompressedFile</span></span><br><span data-ttu-id="3b6ba-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-224">
        - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-225">
        - File</span></span><br><span data-ttu-id="3b6ba-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-226">
        - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-228">
        - Selection</span></span><br><span data-ttu-id="3b6ba-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-229">
        - Settings</span></span><br><span data-ttu-id="3b6ba-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-230">
        - TableBindings</span></span><br><span data-ttu-id="3b6ba-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-231">
        - TableCoercion</span></span><br><span data-ttu-id="3b6ba-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-232">
        - TextBindings</span></span><br><span data-ttu-id="3b6ba-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-234">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="3b6ba-234">Sideload Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="3b6ba-235">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-235">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="3b6ba-236">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-236">- TaskPane</span></span><br><span data-ttu-id="3b6ba-237">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-237">
        - Content</span></span></td>
    <td><span data-ttu-id="3b6ba-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3b6ba-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3b6ba-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3b6ba-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3b6ba-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3b6ba-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3b6ba-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3b6ba-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3b6ba-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3b6ba-249">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-249">- BindingEvents</span></span><br><span data-ttu-id="3b6ba-250">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-250">
        - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-251">
        - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-251">
        - File</span></span><br><span data-ttu-id="3b6ba-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-252">
        - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-253">
        - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-254">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-254">
        - Selection</span></span><br><span data-ttu-id="3b6ba-255">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-255">
        - Settings</span></span><br><span data-ttu-id="3b6ba-256">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-256">
        - TableBindings</span></span><br><span data-ttu-id="3b6ba-257">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-257">
        - TableCoercion</span></span><br><span data-ttu-id="3b6ba-258">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-258">
        - TextBindings</span></span><br><span data-ttu-id="3b6ba-259">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-259">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-260">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="3b6ba-260">Office apps on Mac</span></span><br><span data-ttu-id="3b6ba-261">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-261">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="3b6ba-262">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-262">- TaskPane</span></span><br><span data-ttu-id="3b6ba-263">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-263">
        - Content</span></span><br><span data-ttu-id="3b6ba-264">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="3b6ba-264">
        - Custom Functions</span></span><br><span data-ttu-id="3b6ba-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="3b6ba-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3b6ba-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3b6ba-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3b6ba-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3b6ba-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3b6ba-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3b6ba-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3b6ba-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3b6ba-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3b6ba-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="3b6ba-278">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-278">- BindingEvents</span></span><br><span data-ttu-id="3b6ba-279">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-279">
        - CompressedFile</span></span><br><span data-ttu-id="3b6ba-280">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-280">
        - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-281">
        - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-281">
        - File</span></span><br><span data-ttu-id="3b6ba-282">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-282">
        - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-283">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-283">
        - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-284">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-284">
        - PdfFile</span></span><br><span data-ttu-id="3b6ba-285">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-285">
        - Selection</span></span><br><span data-ttu-id="3b6ba-286">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-286">
        - Settings</span></span><br><span data-ttu-id="3b6ba-287">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-287">
        - TableBindings</span></span><br><span data-ttu-id="3b6ba-288">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-288">
        - TableCoercion</span></span><br><span data-ttu-id="3b6ba-289">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-289">
        - TextBindings</span></span><br><span data-ttu-id="3b6ba-290">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-290">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-291">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="3b6ba-291">Office 2019 for Mac</span></span><br><span data-ttu-id="3b6ba-292">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-292">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3b6ba-293">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-293">- TaskPane</span></span><br><span data-ttu-id="3b6ba-294">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-294">
        - Content</span></span><br><span data-ttu-id="3b6ba-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="3b6ba-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3b6ba-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3b6ba-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3b6ba-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3b6ba-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3b6ba-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3b6ba-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3b6ba-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3b6ba-306">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-306">- BindingEvents</span></span><br><span data-ttu-id="3b6ba-307">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-307">
        - CompressedFile</span></span><br><span data-ttu-id="3b6ba-308">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-308">
        - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-309">
        - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-309">
        - File</span></span><br><span data-ttu-id="3b6ba-310">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-310">
        - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-311">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-311">
        - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-312">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-312">
        - PdfFile</span></span><br><span data-ttu-id="3b6ba-313">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-313">
        - Selection</span></span><br><span data-ttu-id="3b6ba-314">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-314">
        - Settings</span></span><br><span data-ttu-id="3b6ba-315">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-315">
        - TableBindings</span></span><br><span data-ttu-id="3b6ba-316">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-316">
        - TableCoercion</span></span><br><span data-ttu-id="3b6ba-317">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-317">
        - TextBindings</span></span><br><span data-ttu-id="3b6ba-318">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-318">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-319">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="3b6ba-319">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="3b6ba-320">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-320">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3b6ba-321">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-321">- TaskPane</span></span><br><span data-ttu-id="3b6ba-322">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-322">
        - Content</span></span></td>
    <td><span data-ttu-id="3b6ba-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3b6ba-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3b6ba-326">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-326">- BindingEvents</span></span><br><span data-ttu-id="3b6ba-327">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-327">
        - CompressedFile</span></span><br><span data-ttu-id="3b6ba-328">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-328">
        - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-329">
        - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-329">
        - File</span></span><br><span data-ttu-id="3b6ba-330">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-330">
        - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-331">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-331">
        - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-332">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-332">
        - PdfFile</span></span><br><span data-ttu-id="3b6ba-333">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-333">
        - Selection</span></span><br><span data-ttu-id="3b6ba-334">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-334">
        - Settings</span></span><br><span data-ttu-id="3b6ba-335">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-335">
        - TableBindings</span></span><br><span data-ttu-id="3b6ba-336">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-336">
        - TableCoercion</span></span><br><span data-ttu-id="3b6ba-337">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-337">
        - TextBindings</span></span><br><span data-ttu-id="3b6ba-338">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-338">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="3b6ba-339">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-339">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="3b6ba-340">自定义函数</span><span class="sxs-lookup"><span data-stu-id="3b6ba-340">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="3b6ba-341">平台</span><span class="sxs-lookup"><span data-stu-id="3b6ba-341">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="3b6ba-342">扩展点</span><span class="sxs-lookup"><span data-stu-id="3b6ba-342">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="3b6ba-343">API 要求集</span><span class="sxs-lookup"><span data-stu-id="3b6ba-343">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="3b6ba-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-345">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="3b6ba-345">Office on the web</span></span></td>
    <td><span data-ttu-id="3b6ba-346">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="3b6ba-346">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="3b6ba-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-348">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="3b6ba-348">Office on Windows</span></span><br><span data-ttu-id="3b6ba-349">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-349">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="3b6ba-350">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="3b6ba-350">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="3b6ba-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-352">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="3b6ba-352">Office for Mac</span></span><br><span data-ttu-id="3b6ba-353">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-353">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="3b6ba-354">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="3b6ba-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="3b6ba-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="3b6ba-356">Outlook</span><span class="sxs-lookup"><span data-stu-id="3b6ba-356">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3b6ba-357">平台</span><span class="sxs-lookup"><span data-stu-id="3b6ba-357">Platform</span></span></th>
    <th><span data-ttu-id="3b6ba-358">扩展点</span><span class="sxs-lookup"><span data-stu-id="3b6ba-358">Extension points</span></span></th>
    <th><span data-ttu-id="3b6ba-359">API 要求集</span><span class="sxs-lookup"><span data-stu-id="3b6ba-359">API requirement sets</span></span></th>
    <th><span data-ttu-id="3b6ba-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-361">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="3b6ba-361">Office on the web</span></span><br><span data-ttu-id="3b6ba-362">（新式）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-362">modern</span></span></td>
    <td> <span data-ttu-id="3b6ba-363">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="3b6ba-363">- Mail Read</span></span><br><span data-ttu-id="3b6ba-364">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="3b6ba-364">
      - Mail Compose</span></span><br><span data-ttu-id="3b6ba-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3b6ba-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3b6ba-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3b6ba-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3b6ba-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3b6ba-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3b6ba-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="3b6ba-373">不可用</span><span class="sxs-lookup"><span data-stu-id="3b6ba-373">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-374">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="3b6ba-374">Office on the web</span></span><br><span data-ttu-id="3b6ba-375">（经典）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-375">classic</span></span></td>
    <td> <span data-ttu-id="3b6ba-376">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="3b6ba-376">- Mail Read</span></span><br><span data-ttu-id="3b6ba-377">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="3b6ba-377">
      - Mail Compose</span></span><br><span data-ttu-id="3b6ba-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3b6ba-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3b6ba-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3b6ba-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3b6ba-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3b6ba-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="3b6ba-385">不可用</span><span class="sxs-lookup"><span data-stu-id="3b6ba-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-386">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="3b6ba-386">Office on Windows</span></span><br><span data-ttu-id="3b6ba-387">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-387">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3b6ba-388">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="3b6ba-388">- Mail Read</span></span><br><span data-ttu-id="3b6ba-389">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="3b6ba-389">
      - Mail Compose</span></span><br><span data-ttu-id="3b6ba-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="3b6ba-391">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="3b6ba-391">
      - Modules</span></span></td>
    <td> <span data-ttu-id="3b6ba-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3b6ba-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3b6ba-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3b6ba-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3b6ba-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3b6ba-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3b6ba-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="3b6ba-399">不可用</span><span class="sxs-lookup"><span data-stu-id="3b6ba-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-400">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="3b6ba-400">Office 2019 on Windows</span></span><br><span data-ttu-id="3b6ba-401">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-401">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-402">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="3b6ba-402">- Mail Read</span></span><br><span data-ttu-id="3b6ba-403">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="3b6ba-403">
      - Mail Compose</span></span><br><span data-ttu-id="3b6ba-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="3b6ba-405">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="3b6ba-405">
      - Modules</span></span></td>
    <td> <span data-ttu-id="3b6ba-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3b6ba-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3b6ba-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3b6ba-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3b6ba-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3b6ba-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3b6ba-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="3b6ba-413">不可用</span><span class="sxs-lookup"><span data-stu-id="3b6ba-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-414">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="3b6ba-414">Office 2016 on Windows</span></span><br><span data-ttu-id="3b6ba-415">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-416">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="3b6ba-416">- Mail Read</span></span><br><span data-ttu-id="3b6ba-417">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="3b6ba-417">
      - Mail Compose</span></span><br><span data-ttu-id="3b6ba-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="3b6ba-419">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="3b6ba-419">
      - Modules</span></span></td>
    <td> <span data-ttu-id="3b6ba-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3b6ba-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3b6ba-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3b6ba-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="3b6ba-424">不可用</span><span class="sxs-lookup"><span data-stu-id="3b6ba-424">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-425">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="3b6ba-425">Office 2013 on Windows</span></span><br><span data-ttu-id="3b6ba-426">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-426">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-427">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="3b6ba-427">- Mail Read</span></span><br><span data-ttu-id="3b6ba-428">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="3b6ba-428">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="3b6ba-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3b6ba-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3b6ba-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="3b6ba-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="3b6ba-433">不可用</span><span class="sxs-lookup"><span data-stu-id="3b6ba-433">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-434">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="3b6ba-434">Office apps on iOS</span></span><br><span data-ttu-id="3b6ba-435">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-435">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3b6ba-436">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="3b6ba-436">- Mail Read</span></span><br><span data-ttu-id="3b6ba-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3b6ba-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3b6ba-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3b6ba-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3b6ba-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="3b6ba-443">不可用</span><span class="sxs-lookup"><span data-stu-id="3b6ba-443">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-444">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="3b6ba-444">Office apps on Mac</span></span><br><span data-ttu-id="3b6ba-445">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-445">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3b6ba-446">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="3b6ba-446">- Mail Read</span></span><br><span data-ttu-id="3b6ba-447">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="3b6ba-447">
      - Mail Compose</span></span><br><span data-ttu-id="3b6ba-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3b6ba-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3b6ba-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3b6ba-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3b6ba-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3b6ba-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3b6ba-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="3b6ba-456">不可用</span><span class="sxs-lookup"><span data-stu-id="3b6ba-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-457">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="3b6ba-457">Office 2019 for Mac</span></span><br><span data-ttu-id="3b6ba-458">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-458">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-459">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="3b6ba-459">- Mail Read</span></span><br><span data-ttu-id="3b6ba-460">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="3b6ba-460">
      - Mail Compose</span></span><br><span data-ttu-id="3b6ba-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3b6ba-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3b6ba-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3b6ba-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3b6ba-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3b6ba-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="3b6ba-468">不可用</span><span class="sxs-lookup"><span data-stu-id="3b6ba-468">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-469">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="3b6ba-469">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="3b6ba-470">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-470">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-471">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="3b6ba-471">- Mail Read</span></span><br><span data-ttu-id="3b6ba-472">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="3b6ba-472">
      - Mail Compose</span></span><br><span data-ttu-id="3b6ba-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3b6ba-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3b6ba-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3b6ba-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3b6ba-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3b6ba-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="3b6ba-480">不可用</span><span class="sxs-lookup"><span data-stu-id="3b6ba-480">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-481">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="3b6ba-481">Office apps on Android</span></span><br><span data-ttu-id="3b6ba-482">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-482">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3b6ba-483">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="3b6ba-483">- Mail Read</span></span><br><span data-ttu-id="3b6ba-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3b6ba-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3b6ba-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3b6ba-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3b6ba-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="3b6ba-490">不可用</span><span class="sxs-lookup"><span data-stu-id="3b6ba-490">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="3b6ba-491">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-491">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3b6ba-492">要求集的客户端支持可能受到 Exchange 服务器支持的限制。</span><span class="sxs-lookup"><span data-stu-id="3b6ba-492">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="3b6ba-493">有关 Exchange 服务器和 Outlook 客户端支持的要求集范围的详细信息，请参阅 [Outlook JavaScript API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="3b6ba-493">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="3b6ba-494">Word</span><span class="sxs-lookup"><span data-stu-id="3b6ba-494">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3b6ba-495">平台</span><span class="sxs-lookup"><span data-stu-id="3b6ba-495">Platform</span></span></th>
    <th><span data-ttu-id="3b6ba-496">扩展点</span><span class="sxs-lookup"><span data-stu-id="3b6ba-496">Extension points</span></span></th>
    <th><span data-ttu-id="3b6ba-497">API 要求集</span><span class="sxs-lookup"><span data-stu-id="3b6ba-497">API requirement sets</span></span></th>
    <th><span data-ttu-id="3b6ba-498"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-498"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-499">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="3b6ba-499">Office on the web</span></span></td>
    <td> <span data-ttu-id="3b6ba-500">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-500">- TaskPane</span></span><br><span data-ttu-id="3b6ba-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-502">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-502">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3b6ba-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3b6ba-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3b6ba-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-508">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-508">- BindingEvents</span></span><br><span data-ttu-id="3b6ba-509">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3b6ba-509">
         - CustomXmlParts</span></span><br><span data-ttu-id="3b6ba-510">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-510">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-511">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-511">
         - File</span></span><br><span data-ttu-id="3b6ba-512">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-512">
         - HtmlCoercion</span></span><br><span data-ttu-id="3b6ba-513">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-513">
         - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-514">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-514">
         - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-515">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-515">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3b6ba-516">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-516">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-517">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-517">
         - Selection</span></span><br><span data-ttu-id="3b6ba-518">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-518">
         - Settings</span></span><br><span data-ttu-id="3b6ba-519">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-519">
         - TableBindings</span></span><br><span data-ttu-id="3b6ba-520">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-520">
         - TableCoercion</span></span><br><span data-ttu-id="3b6ba-521">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-521">
         - TextBindings</span></span><br><span data-ttu-id="3b6ba-522">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-522">
         - TextCoercion</span></span><br><span data-ttu-id="3b6ba-523">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-523">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-524">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="3b6ba-524">Office on Windows</span></span><br><span data-ttu-id="3b6ba-525">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-525">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3b6ba-526">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-526">- TaskPane</span></span><br><span data-ttu-id="3b6ba-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3b6ba-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3b6ba-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3b6ba-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-534">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-534">- BindingEvents</span></span><br><span data-ttu-id="3b6ba-535">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-535">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-536">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3b6ba-536">
         - CustomXmlParts</span></span><br><span data-ttu-id="3b6ba-537">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-537">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-538">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-538">
         - File</span></span><br><span data-ttu-id="3b6ba-539">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-539">
         - HtmlCoercion</span></span><br><span data-ttu-id="3b6ba-540">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-540">
         - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-541">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-541">
         - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-542">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-542">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3b6ba-543">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-543">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-544">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-544">
         - Selection</span></span><br><span data-ttu-id="3b6ba-545">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-545">
         - Settings</span></span><br><span data-ttu-id="3b6ba-546">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-546">
         - TableBindings</span></span><br><span data-ttu-id="3b6ba-547">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-547">
         - TableCoercion</span></span><br><span data-ttu-id="3b6ba-548">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-548">
         - TextBindings</span></span><br><span data-ttu-id="3b6ba-549">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-549">
         - TextCoercion</span></span><br><span data-ttu-id="3b6ba-550">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-550">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-551">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="3b6ba-551">Office 2019 on Windows</span></span><br><span data-ttu-id="3b6ba-552">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-552">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-553">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-553">- TaskPane</span></span><br><span data-ttu-id="3b6ba-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3b6ba-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3b6ba-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-560">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-560">- BindingEvents</span></span><br><span data-ttu-id="3b6ba-561">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-561">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-562">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3b6ba-562">
         - CustomXmlParts</span></span><br><span data-ttu-id="3b6ba-563">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-563">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-564">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-564">
         - File</span></span><br><span data-ttu-id="3b6ba-565">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-565">
         - HtmlCoercion</span></span><br><span data-ttu-id="3b6ba-566">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-566">
         - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-567">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-567">
         - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-568">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-568">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3b6ba-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-569">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-570">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-570">
         - Selection</span></span><br><span data-ttu-id="3b6ba-571">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-571">
         - Settings</span></span><br><span data-ttu-id="3b6ba-572">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-572">
         - TableBindings</span></span><br><span data-ttu-id="3b6ba-573">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-573">
         - TableCoercion</span></span><br><span data-ttu-id="3b6ba-574">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-574">
         - TextBindings</span></span><br><span data-ttu-id="3b6ba-575">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-575">
         - TextCoercion</span></span><br><span data-ttu-id="3b6ba-576">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-576">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-577">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="3b6ba-577">Office 2016 on Windows</span></span><br><span data-ttu-id="3b6ba-578">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-578">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-579">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-579">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3b6ba-580">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-580">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3b6ba-582">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-582">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-583">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-583">- BindingEvents</span></span><br><span data-ttu-id="3b6ba-584">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-584">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-585">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3b6ba-585">
         - CustomXmlParts</span></span><br><span data-ttu-id="3b6ba-586">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-586">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-587">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-587">
         - File</span></span><br><span data-ttu-id="3b6ba-588">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-588">
         - HtmlCoercion</span></span><br><span data-ttu-id="3b6ba-589">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-589">
         - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-590">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-590">
         - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-591">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-591">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3b6ba-592">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-592">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-593">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-593">
         - Selection</span></span><br><span data-ttu-id="3b6ba-594">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-594">
         - Settings</span></span><br><span data-ttu-id="3b6ba-595">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-595">
         - TableBindings</span></span><br><span data-ttu-id="3b6ba-596">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-596">
         - TableCoercion</span></span><br><span data-ttu-id="3b6ba-597">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-597">
         - TextBindings</span></span><br><span data-ttu-id="3b6ba-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-598">
         - TextCoercion</span></span><br><span data-ttu-id="3b6ba-599">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-599">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-600">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="3b6ba-600">Office 2013 on Windows</span></span><br><span data-ttu-id="3b6ba-601">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-601">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-602">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-602">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3b6ba-603">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-603">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="3b6ba-604">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-604">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-605">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-605">- BindingEvents</span></span><br><span data-ttu-id="3b6ba-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-606">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-607">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3b6ba-607">
         - CustomXmlParts</span></span><br><span data-ttu-id="3b6ba-608">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-608">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-609">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-609">
         - File</span></span><br><span data-ttu-id="3b6ba-610">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-610">
         - HtmlCoercion</span></span><br><span data-ttu-id="3b6ba-611">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-611">
         - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-612">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-612">
         - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-613">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-613">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3b6ba-614">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-614">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-615">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-615">
         - Selection</span></span><br><span data-ttu-id="3b6ba-616">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-616">
         - Settings</span></span><br><span data-ttu-id="3b6ba-617">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-617">
         - TableBindings</span></span><br><span data-ttu-id="3b6ba-618">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-618">
         - TableCoercion</span></span><br><span data-ttu-id="3b6ba-619">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-619">
         - TextBindings</span></span><br><span data-ttu-id="3b6ba-620">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-620">
         - TextCoercion</span></span><br><span data-ttu-id="3b6ba-621">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-621">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-622">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="3b6ba-622">Sideload Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="3b6ba-623">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-623">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3b6ba-624">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-624">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3b6ba-625">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-625">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3b6ba-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3b6ba-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="3b6ba-630">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-630">- BindingEvents</span></span><br><span data-ttu-id="3b6ba-631">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-631">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-632">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3b6ba-632">
         - CustomXmlParts</span></span><br><span data-ttu-id="3b6ba-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-633">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-634">
         - File</span></span><br><span data-ttu-id="3b6ba-635">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-635">
         - HtmlCoercion</span></span><br><span data-ttu-id="3b6ba-636">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-636">
         - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-637">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-637">
         - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-638">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-638">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3b6ba-639">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-639">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-640">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-640">
         - Selection</span></span><br><span data-ttu-id="3b6ba-641">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-641">
         - Settings</span></span><br><span data-ttu-id="3b6ba-642">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-642">
         - TableBindings</span></span><br><span data-ttu-id="3b6ba-643">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-643">
         - TableCoercion</span></span><br><span data-ttu-id="3b6ba-644">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-644">
         - TextBindings</span></span><br><span data-ttu-id="3b6ba-645">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-645">
         - TextCoercion</span></span><br><span data-ttu-id="3b6ba-646">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-646">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-647">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="3b6ba-647">Office apps on Mac</span></span><br><span data-ttu-id="3b6ba-648">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-648">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3b6ba-649">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-649">- TaskPane</span></span><br><span data-ttu-id="3b6ba-650">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-650">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3b6ba-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3b6ba-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3b6ba-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="3b6ba-657">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-657">- BindingEvents</span></span><br><span data-ttu-id="3b6ba-658">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-658">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-659">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3b6ba-659">
         - CustomXmlParts</span></span><br><span data-ttu-id="3b6ba-660">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-660">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-661">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-661">
         - File</span></span><br><span data-ttu-id="3b6ba-662">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-662">
         - HtmlCoercion</span></span><br><span data-ttu-id="3b6ba-663">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-663">
         - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-664">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-664">
         - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-665">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-665">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3b6ba-666">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-666">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-667">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-667">
         - Selection</span></span><br><span data-ttu-id="3b6ba-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-668">
         - Settings</span></span><br><span data-ttu-id="3b6ba-669">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-669">
         - TableBindings</span></span><br><span data-ttu-id="3b6ba-670">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-670">
         - TableCoercion</span></span><br><span data-ttu-id="3b6ba-671">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-671">
         - TextBindings</span></span><br><span data-ttu-id="3b6ba-672">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-672">
         - TextCoercion</span></span><br><span data-ttu-id="3b6ba-673">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-673">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-674">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="3b6ba-674">Office 2019 for Mac</span></span><br><span data-ttu-id="3b6ba-675">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-675">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-676">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-676">- TaskPane</span></span><br><span data-ttu-id="3b6ba-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3b6ba-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3b6ba-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="3b6ba-683">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-683">- BindingEvents</span></span><br><span data-ttu-id="3b6ba-684">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-684">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-685">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3b6ba-685">
         - CustomXmlParts</span></span><br><span data-ttu-id="3b6ba-686">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-686">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-687">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-687">
         - File</span></span><br><span data-ttu-id="3b6ba-688">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-688">
         - HtmlCoercion</span></span><br><span data-ttu-id="3b6ba-689">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-689">
         - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-690">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-690">
         - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-691">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-691">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3b6ba-692">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-692">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-693">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-693">
         - Selection</span></span><br><span data-ttu-id="3b6ba-694">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-694">
         - Settings</span></span><br><span data-ttu-id="3b6ba-695">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-695">
         - TableBindings</span></span><br><span data-ttu-id="3b6ba-696">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-696">
         - TableCoercion</span></span><br><span data-ttu-id="3b6ba-697">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-697">
         - TextBindings</span></span><br><span data-ttu-id="3b6ba-698">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-698">
         - TextCoercion</span></span><br><span data-ttu-id="3b6ba-699">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-699">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-700">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="3b6ba-700">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="3b6ba-701">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-701">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-702">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-702">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3b6ba-703">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-703">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-704">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-704">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3b6ba-705">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-705">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-706">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-706">- BindingEvents</span></span><br><span data-ttu-id="3b6ba-707">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-707">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-708">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3b6ba-708">
         - CustomXmlParts</span></span><br><span data-ttu-id="3b6ba-709">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-709">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-710">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-710">
         - File</span></span><br><span data-ttu-id="3b6ba-711">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-711">
         - HtmlCoercion</span></span><br><span data-ttu-id="3b6ba-712">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-712">
         - MatrixBindings</span></span><br><span data-ttu-id="3b6ba-713">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-713">
         - MatrixCoercion</span></span><br><span data-ttu-id="3b6ba-714">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-714">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3b6ba-715">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-715">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-716">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-716">
         - Selection</span></span><br><span data-ttu-id="3b6ba-717">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-717">
         - Settings</span></span><br><span data-ttu-id="3b6ba-718">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-718">
         - TableBindings</span></span><br><span data-ttu-id="3b6ba-719">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-719">
         - TableCoercion</span></span><br><span data-ttu-id="3b6ba-720">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-720">
         - TextBindings</span></span><br><span data-ttu-id="3b6ba-721">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-721">
         - TextCoercion</span></span><br><span data-ttu-id="3b6ba-722">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-722">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="3b6ba-723">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-723">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="3b6ba-724">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="3b6ba-724">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3b6ba-725">平台</span><span class="sxs-lookup"><span data-stu-id="3b6ba-725">Platform</span></span></th>
    <th><span data-ttu-id="3b6ba-726">扩展点</span><span class="sxs-lookup"><span data-stu-id="3b6ba-726">Extension points</span></span></th>
    <th><span data-ttu-id="3b6ba-727">API 要求集</span><span class="sxs-lookup"><span data-stu-id="3b6ba-727">API requirement sets</span></span></th>
    <th><span data-ttu-id="3b6ba-728"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-728"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-729">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="3b6ba-729">Office on the web</span></span></td>
    <td> <span data-ttu-id="3b6ba-730">- 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-730">- Content</span></span><br><span data-ttu-id="3b6ba-731">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-731">
         - TaskPane</span></span><br><span data-ttu-id="3b6ba-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-733">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-733">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3b6ba-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3b6ba-737">- ActiveView</span></span><br><span data-ttu-id="3b6ba-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-738">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-739">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-740">
         - File</span></span><br><span data-ttu-id="3b6ba-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-741">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-742">
         - Selection</span></span><br><span data-ttu-id="3b6ba-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-743">
         - Settings</span></span><br><span data-ttu-id="3b6ba-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-745">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="3b6ba-745">Office on Windows</span></span><br><span data-ttu-id="3b6ba-746">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-746">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3b6ba-747">- 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-747">- Content</span></span><br><span data-ttu-id="3b6ba-748">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-748">
         - TaskPane</span></span><br><span data-ttu-id="3b6ba-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-750">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-750">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3b6ba-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-754">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3b6ba-754">- ActiveView</span></span><br><span data-ttu-id="3b6ba-755">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-755">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-756">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-756">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-757">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-757">
         - File</span></span><br><span data-ttu-id="3b6ba-758">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-758">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-759">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-759">
         - Selection</span></span><br><span data-ttu-id="3b6ba-760">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-760">
         - Settings</span></span><br><span data-ttu-id="3b6ba-761">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-761">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-762">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="3b6ba-762">Office 2019 on Windows</span></span><br><span data-ttu-id="3b6ba-763">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-763">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-764">- 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-764">- Content</span></span><br><span data-ttu-id="3b6ba-765">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-765">
         - TaskPane</span></span><br><span data-ttu-id="3b6ba-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-767">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-767">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-768">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-768">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-769">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3b6ba-769">- ActiveView</span></span><br><span data-ttu-id="3b6ba-770">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-770">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-771">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-771">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-772">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-772">
         - File</span></span><br><span data-ttu-id="3b6ba-773">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-773">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-774">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-774">
         - Selection</span></span><br><span data-ttu-id="3b6ba-775">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-775">
         - Settings</span></span><br><span data-ttu-id="3b6ba-776">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-776">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-777">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="3b6ba-777">Office 2016 on Windows</span></span><br><span data-ttu-id="3b6ba-778">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-778">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-779">- 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-779">- Content</span></span><br><span data-ttu-id="3b6ba-780">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-780">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="3b6ba-781">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-781">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="3b6ba-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-783">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3b6ba-783">- ActiveView</span></span><br><span data-ttu-id="3b6ba-784">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-784">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-785">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-785">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-786">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-786">
         - File</span></span><br><span data-ttu-id="3b6ba-787">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-787">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-788">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-788">
         - Selection</span></span><br><span data-ttu-id="3b6ba-789">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-789">
         - Settings</span></span><br><span data-ttu-id="3b6ba-790">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-790">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-791">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="3b6ba-791">Office 2013 on Windows</span></span><br><span data-ttu-id="3b6ba-792">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-792">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-793">- 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-793">- Content</span></span><br><span data-ttu-id="3b6ba-794">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-794">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="3b6ba-795">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-795">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="3b6ba-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-797">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3b6ba-797">- ActiveView</span></span><br><span data-ttu-id="3b6ba-798">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-798">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-799">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-799">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-800">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-800">
         - File</span></span><br><span data-ttu-id="3b6ba-801">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-801">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-802">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-802">
         - Selection</span></span><br><span data-ttu-id="3b6ba-803">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-803">
         - Settings</span></span><br><span data-ttu-id="3b6ba-804">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-804">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-805">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="3b6ba-805">Sideload Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="3b6ba-806">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-806">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3b6ba-807">- 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-807">- Content</span></span><br><span data-ttu-id="3b6ba-808">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-808">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="3b6ba-809">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-809">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-812">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3b6ba-812">- ActiveView</span></span><br><span data-ttu-id="3b6ba-813">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-813">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-814">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-814">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-815">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-815">
         - File</span></span><br><span data-ttu-id="3b6ba-816">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-816">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-817">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-817">
         - Selection</span></span><br><span data-ttu-id="3b6ba-818">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-818">
         - Settings</span></span><br><span data-ttu-id="3b6ba-819">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-819">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-820">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="3b6ba-820">Office apps on Mac</span></span><br><span data-ttu-id="3b6ba-821">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-821">Outlook on Mac (connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3b6ba-822">- 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-822">- Content</span></span><br><span data-ttu-id="3b6ba-823">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-823">
         - TaskPane</span></span><br><span data-ttu-id="3b6ba-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-825">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-825">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3b6ba-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-829">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3b6ba-829">- ActiveView</span></span><br><span data-ttu-id="3b6ba-830">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-830">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-831">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-831">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-832">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-832">
         - File</span></span><br><span data-ttu-id="3b6ba-833">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-833">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-834">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-834">
         - Selection</span></span><br><span data-ttu-id="3b6ba-835">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-835">
         - Settings</span></span><br><span data-ttu-id="3b6ba-836">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-836">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-837">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="3b6ba-837">Office 2019 for Mac</span></span><br><span data-ttu-id="3b6ba-838">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-838">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-839">- 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-839">- Content</span></span><br><span data-ttu-id="3b6ba-840">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-840">
         - TaskPane</span></span><br><span data-ttu-id="3b6ba-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-843">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-843">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-844">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3b6ba-844">- ActiveView</span></span><br><span data-ttu-id="3b6ba-845">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-845">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-846">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-846">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-847">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-847">
         - File</span></span><br><span data-ttu-id="3b6ba-848">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-848">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-849">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-849">
         - Selection</span></span><br><span data-ttu-id="3b6ba-850">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-850">
         - Settings</span></span><br><span data-ttu-id="3b6ba-851">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-851">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-852">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="3b6ba-852">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="3b6ba-853">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-853">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-854">- 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-854">- Content</span></span><br><span data-ttu-id="3b6ba-855">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-855">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="3b6ba-856">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-856">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="3b6ba-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-858">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3b6ba-858">- ActiveView</span></span><br><span data-ttu-id="3b6ba-859">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-859">
         - CompressedFile</span></span><br><span data-ttu-id="3b6ba-860">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-860">
         - DocumentEvents</span></span><br><span data-ttu-id="3b6ba-861">
         - File</span><span class="sxs-lookup"><span data-stu-id="3b6ba-861">
         - File</span></span><br><span data-ttu-id="3b6ba-862">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3b6ba-862">
         - PdfFile</span></span><br><span data-ttu-id="3b6ba-863">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-863">
         - Selection</span></span><br><span data-ttu-id="3b6ba-864">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-864">
         - Settings</span></span><br><span data-ttu-id="3b6ba-865">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-865">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="3b6ba-866">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="3b6ba-866">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="3b6ba-867">OneNote</span><span class="sxs-lookup"><span data-stu-id="3b6ba-867">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3b6ba-868">平台</span><span class="sxs-lookup"><span data-stu-id="3b6ba-868">Platform</span></span></th>
    <th><span data-ttu-id="3b6ba-869">扩展点</span><span class="sxs-lookup"><span data-stu-id="3b6ba-869">Extension points</span></span></th>
    <th><span data-ttu-id="3b6ba-870">API 要求集</span><span class="sxs-lookup"><span data-stu-id="3b6ba-870">API requirement sets</span></span></th>
    <th><span data-ttu-id="3b6ba-871"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-871"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-872">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="3b6ba-872">Office on the web</span></span></td>
    <td> <span data-ttu-id="3b6ba-873">- 内容</span><span class="sxs-lookup"><span data-stu-id="3b6ba-873">- Content</span></span><br><span data-ttu-id="3b6ba-874">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-874">
         - TaskPane</span></span><br><span data-ttu-id="3b6ba-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-876">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-876">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3b6ba-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-879">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3b6ba-879">- DocumentEvents</span></span><br><span data-ttu-id="3b6ba-880">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-880">
         - HtmlCoercion</span></span><br><span data-ttu-id="3b6ba-881">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3b6ba-881">
         - Settings</span></span><br><span data-ttu-id="3b6ba-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-882">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="3b6ba-883">项目</span><span class="sxs-lookup"><span data-stu-id="3b6ba-883">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3b6ba-884">平台</span><span class="sxs-lookup"><span data-stu-id="3b6ba-884">Platform</span></span></th>
    <th><span data-ttu-id="3b6ba-885">扩展点</span><span class="sxs-lookup"><span data-stu-id="3b6ba-885">Extension points</span></span></th>
    <th><span data-ttu-id="3b6ba-886">API 要求集</span><span class="sxs-lookup"><span data-stu-id="3b6ba-886">API requirement sets</span></span></th>
    <th><span data-ttu-id="3b6ba-887"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-887"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-888">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="3b6ba-888">Office 2019 on Windows</span></span><br><span data-ttu-id="3b6ba-889">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-889">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-890">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-890">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3b6ba-891">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-891">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-892">- Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-892">- Selection</span></span><br><span data-ttu-id="3b6ba-893">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-893">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-894">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="3b6ba-894">Office 2016 on Windows</span></span><br><span data-ttu-id="3b6ba-895">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-895">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-896">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-896">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3b6ba-897">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-897">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-898">- Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-898">- Selection</span></span><br><span data-ttu-id="3b6ba-899">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-899">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3b6ba-900">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="3b6ba-900">Office 2013 on Windows</span></span><br><span data-ttu-id="3b6ba-901">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-901">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3b6ba-902">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="3b6ba-902">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3b6ba-903">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3b6ba-903">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="3b6ba-904">- Selection</span><span class="sxs-lookup"><span data-stu-id="3b6ba-904">- Selection</span></span><br><span data-ttu-id="3b6ba-905">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3b6ba-905">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="3b6ba-906">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3b6ba-906">See also</span></span>

- [<span data-ttu-id="3b6ba-907">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="3b6ba-907">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="3b6ba-908">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="3b6ba-908">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="3b6ba-909">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="3b6ba-909">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="3b6ba-910">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="3b6ba-910">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="3b6ba-911">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="3b6ba-911">JavaScript API for Office reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="3b6ba-912">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="3b6ba-912">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="3b6ba-913">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-913">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="3b6ba-914">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="3b6ba-914">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="3b6ba-915">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="3b6ba-915">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="3b6ba-916">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="3b6ba-916">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="3b6ba-917">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="3b6ba-917">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)

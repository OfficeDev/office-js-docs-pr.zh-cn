---
title: Office 外接程序主机和平台可用性
description: Excel、Word、Outlook、PowerPoint、OneNote 和项目支持的要求集。
ms.date: 04/03/2019
localization_priority: Priority
ms.openlocfilehash: a9ecd44edf9221a403eb42756cd1e9f5e676ad01
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448145"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="0b7f8-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="0b7f8-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="0b7f8-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="0b7f8-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="0b7f8-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="0b7f8-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="0b7f8-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="0b7f8-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="0b7f8-108">Excel</span><span class="sxs-lookup"><span data-stu-id="0b7f8-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="0b7f8-109">平台</span><span class="sxs-lookup"><span data-stu-id="0b7f8-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="0b7f8-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="0b7f8-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="0b7f8-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="0b7f8-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="0b7f8-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="0b7f8-113">Office Online</span></span></td>
    <td> <span data-ttu-id="0b7f8-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-114">- TaskPane</span></span><br><span data-ttu-id="0b7f8-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-115">
        - Content</span></span><br><span data-ttu-id="0b7f8-116">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="0b7f8-116">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="0b7f8-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0b7f8-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0b7f8-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0b7f8-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0b7f8-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0b7f8-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0b7f8-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0b7f8-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0b7f8-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-126">
        - BindingEvents</span></span><br><span data-ttu-id="0b7f8-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-127">
        - CompressedFile</span></span><br><span data-ttu-id="0b7f8-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-128">
        - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-129">
        - File</span></span><br><span data-ttu-id="0b7f8-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-130">
        - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-132">
        - Selection</span></span><br><span data-ttu-id="0b7f8-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-133">
        - Settings</span></span><br><span data-ttu-id="0b7f8-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-134">
        - TableBindings</span></span><br><span data-ttu-id="0b7f8-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-135">
        - TableCoercion</span></span><br><span data-ttu-id="0b7f8-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-136">
        - TextBindings</span></span><br><span data-ttu-id="0b7f8-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-138">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-138">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-139">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-139">- TaskPane</span></span><br><span data-ttu-id="0b7f8-140">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-140">
        - Content</span></span><br><span data-ttu-id="0b7f8-141">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="0b7f8-141">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="0b7f8-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0b7f8-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0b7f8-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0b7f8-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0b7f8-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0b7f8-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0b7f8-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0b7f8-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0b7f8-151">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-151">
        - BindingEvents</span></span><br><span data-ttu-id="0b7f8-152">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-152">
        - CompressedFile</span></span><br><span data-ttu-id="0b7f8-153">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-153">
        - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-154">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-154">
        - File</span></span><br><span data-ttu-id="0b7f8-155">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-155">
        - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-156">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-156">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-157">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-157">
        - Selection</span></span><br><span data-ttu-id="0b7f8-158">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-158">
        - Settings</span></span><br><span data-ttu-id="0b7f8-159">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-159">
        - TableBindings</span></span><br><span data-ttu-id="0b7f8-160">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-160">
        - TableCoercion</span></span><br><span data-ttu-id="0b7f8-161">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-161">
        - TextBindings</span></span><br><span data-ttu-id="0b7f8-162">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-162">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-163">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-163">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="0b7f8-164">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-164">- TaskPane</span></span><br><span data-ttu-id="0b7f8-165">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-165">
        - Content</span></span><br><span data-ttu-id="0b7f8-166">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-166">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="0b7f8-167">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-167">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-168">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-168">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0b7f8-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0b7f8-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0b7f8-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0b7f8-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0b7f8-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0b7f8-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0b7f8-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0b7f8-176">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-176">- BindingEvents</span></span><br><span data-ttu-id="0b7f8-177">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-177">
        - CompressedFile</span></span><br><span data-ttu-id="0b7f8-178">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-178">
        - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-179">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-179">
        - File</span></span><br><span data-ttu-id="0b7f8-180">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-180">
        - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-181">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-181">
        - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-182">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-182">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-183">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-183">
        - Selection</span></span><br><span data-ttu-id="0b7f8-184">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-184">
        - Settings</span></span><br><span data-ttu-id="0b7f8-185">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-185">
        - TableBindings</span></span><br><span data-ttu-id="0b7f8-186">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-186">
        - TableCoercion</span></span><br><span data-ttu-id="0b7f8-187">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-187">
        - TextBindings</span></span><br><span data-ttu-id="0b7f8-188">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-188">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-189">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-189">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="0b7f8-190">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-190">- TaskPane</span></span><br><span data-ttu-id="0b7f8-191">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-191">
        - Content</span></span></td>
    <td><span data-ttu-id="0b7f8-192">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-192">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-193">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-193">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="0b7f8-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-194">- BindingEvents</span></span><br><span data-ttu-id="0b7f8-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-195">
        - CompressedFile</span></span><br><span data-ttu-id="0b7f8-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-196">
        - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-197">
        - File</span></span><br><span data-ttu-id="0b7f8-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-198">
        - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-199">
        - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-201">
        - Selection</span></span><br><span data-ttu-id="0b7f8-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-202">
        - Settings</span></span><br><span data-ttu-id="0b7f8-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-203">
        - TableBindings</span></span><br><span data-ttu-id="0b7f8-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-204">
        - TableCoercion</span></span><br><span data-ttu-id="0b7f8-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-205">
        - TextBindings</span></span><br><span data-ttu-id="0b7f8-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-207">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-207">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="0b7f8-208">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-208">
        - TaskPane</span></span><br><span data-ttu-id="0b7f8-209">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-209">
        - Content</span></span></td>
    <td>  <span data-ttu-id="0b7f8-210">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-210">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="0b7f8-211">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-211">
        - BindingEvents</span></span><br><span data-ttu-id="0b7f8-212">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-212">
        - CompressedFile</span></span><br><span data-ttu-id="0b7f8-213">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-213">
        - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-214">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-214">
        - File</span></span><br><span data-ttu-id="0b7f8-215">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-215">
        - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-216">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-216">
        - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-217">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-217">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-218">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-218">
        - Selection</span></span><br><span data-ttu-id="0b7f8-219">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-219">
        - Settings</span></span><br><span data-ttu-id="0b7f8-220">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-220">
        - TableBindings</span></span><br><span data-ttu-id="0b7f8-221">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-221">
        - TableCoercion</span></span><br><span data-ttu-id="0b7f8-222">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-222">
        - TextBindings</span></span><br><span data-ttu-id="0b7f8-223">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-223">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-224">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="0b7f8-224">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="0b7f8-225">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-225">- TaskPane</span></span><br><span data-ttu-id="0b7f8-226">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-226">
        - Content</span></span></td>
    <td><span data-ttu-id="0b7f8-227">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-227">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-228">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-228">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0b7f8-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0b7f8-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0b7f8-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0b7f8-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0b7f8-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0b7f8-234">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-234">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0b7f8-235">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-235">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0b7f8-236">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-236">- BindingEvents</span></span><br><span data-ttu-id="0b7f8-237">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-237">
        - CompressedFile</span></span><br><span data-ttu-id="0b7f8-238">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-238">
        - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-239">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-239">
        - File</span></span><br><span data-ttu-id="0b7f8-240">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-240">
        - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-241">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-241">
        - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-242">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-242">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-243">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-243">
        - Selection</span></span><br><span data-ttu-id="0b7f8-244">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-244">
        - Settings</span></span><br><span data-ttu-id="0b7f8-245">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-245">
        - TableBindings</span></span><br><span data-ttu-id="0b7f8-246">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-246">
        - TableCoercion</span></span><br><span data-ttu-id="0b7f8-247">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-247">
        - TextBindings</span></span><br><span data-ttu-id="0b7f8-248">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-248">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-249">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="0b7f8-249">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="0b7f8-250">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-250">- TaskPane</span></span><br><span data-ttu-id="0b7f8-251">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-251">
        - Content</span></span><br><span data-ttu-id="0b7f8-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="0b7f8-253">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-253">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0b7f8-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0b7f8-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0b7f8-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0b7f8-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0b7f8-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0b7f8-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0b7f8-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0b7f8-262">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-262">- BindingEvents</span></span><br><span data-ttu-id="0b7f8-263">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-263">
        - CompressedFile</span></span><br><span data-ttu-id="0b7f8-264">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-264">
        - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-265">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-265">
        - File</span></span><br><span data-ttu-id="0b7f8-266">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-266">
        - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-267">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-267">
        - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-268">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-268">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-269">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-269">
        - PdfFile</span></span><br><span data-ttu-id="0b7f8-270">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-270">
        - Selection</span></span><br><span data-ttu-id="0b7f8-271">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-271">
        - Settings</span></span><br><span data-ttu-id="0b7f8-272">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-272">
        - TableBindings</span></span><br><span data-ttu-id="0b7f8-273">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-273">
        - TableCoercion</span></span><br><span data-ttu-id="0b7f8-274">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-274">
        - TextBindings</span></span><br><span data-ttu-id="0b7f8-275">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-275">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-276">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="0b7f8-276">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="0b7f8-277">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-277">- TaskPane</span></span><br><span data-ttu-id="0b7f8-278">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-278">
        - Content</span></span><br><span data-ttu-id="0b7f8-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="0b7f8-280">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-280">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0b7f8-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0b7f8-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0b7f8-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0b7f8-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0b7f8-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0b7f8-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0b7f8-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0b7f8-289">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-289">- BindingEvents</span></span><br><span data-ttu-id="0b7f8-290">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-290">
        - CompressedFile</span></span><br><span data-ttu-id="0b7f8-291">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-291">
        - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-292">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-292">
        - File</span></span><br><span data-ttu-id="0b7f8-293">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-293">
        - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-294">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-294">
        - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-295">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-295">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-296">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-296">
        - PdfFile</span></span><br><span data-ttu-id="0b7f8-297">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-297">
        - Selection</span></span><br><span data-ttu-id="0b7f8-298">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-298">
        - Settings</span></span><br><span data-ttu-id="0b7f8-299">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-299">
        - TableBindings</span></span><br><span data-ttu-id="0b7f8-300">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-300">
        - TableCoercion</span></span><br><span data-ttu-id="0b7f8-301">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-301">
        - TextBindings</span></span><br><span data-ttu-id="0b7f8-302">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-302">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-303">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="0b7f8-303">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="0b7f8-304">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-304">- TaskPane</span></span><br><span data-ttu-id="0b7f8-305">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-305">
        - Content</span></span></td>
    <td><span data-ttu-id="0b7f8-306">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-306">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="0b7f8-308">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-308">- BindingEvents</span></span><br><span data-ttu-id="0b7f8-309">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-309">
        - CompressedFile</span></span><br><span data-ttu-id="0b7f8-310">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-310">
        - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-311">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-311">
        - File</span></span><br><span data-ttu-id="0b7f8-312">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-312">
        - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-313">
        - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-315">
        - PdfFile</span></span><br><span data-ttu-id="0b7f8-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-316">
        - Selection</span></span><br><span data-ttu-id="0b7f8-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-317">
        - Settings</span></span><br><span data-ttu-id="0b7f8-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-318">
        - TableBindings</span></span><br><span data-ttu-id="0b7f8-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-319">
        - TableCoercion</span></span><br><span data-ttu-id="0b7f8-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-320">
        - TextBindings</span></span><br><span data-ttu-id="0b7f8-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-321">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="0b7f8-322">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-322">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="0b7f8-323">Outlook</span><span class="sxs-lookup"><span data-stu-id="0b7f8-323">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="0b7f8-324">平台</span><span class="sxs-lookup"><span data-stu-id="0b7f8-324">Platform</span></span></th>
    <th><span data-ttu-id="0b7f8-325">扩展点</span><span class="sxs-lookup"><span data-stu-id="0b7f8-325">Extension points</span></span></th>
    <th><span data-ttu-id="0b7f8-326">API 要求集</span><span class="sxs-lookup"><span data-stu-id="0b7f8-326">API requirement sets</span></span></th>
    <th><span data-ttu-id="0b7f8-327"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-327"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-328">Office Online</span><span class="sxs-lookup"><span data-stu-id="0b7f8-328">Office Online</span></span></td>
    <td> <span data-ttu-id="0b7f8-329">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="0b7f8-329">- Mail Read</span></span><br><span data-ttu-id="0b7f8-330">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="0b7f8-330">
      - Mail Compose</span></span><br><span data-ttu-id="0b7f8-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-332">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-332">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b7f8-333">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-333">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b7f8-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b7f8-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b7f8-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0b7f8-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="0b7f8-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="0b7f8-339">不可用</span><span class="sxs-lookup"><span data-stu-id="0b7f8-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-340">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-340">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-341">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="0b7f8-341">- Mail Read</span></span><br><span data-ttu-id="0b7f8-342">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="0b7f8-342">
      - Mail Compose</span></span><br><span data-ttu-id="0b7f8-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="0b7f8-344">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="0b7f8-344">
      - Modules</span></span></td>
    <td> <span data-ttu-id="0b7f8-345">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-345">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b7f8-346">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-346">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b7f8-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b7f8-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b7f8-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0b7f8-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="0b7f8-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="0b7f8-352">不可用</span><span class="sxs-lookup"><span data-stu-id="0b7f8-352">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-353">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-353">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-354">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="0b7f8-354">- Mail Read</span></span><br><span data-ttu-id="0b7f8-355">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="0b7f8-355">
      - Mail Compose</span></span><br><span data-ttu-id="0b7f8-356">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-356">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="0b7f8-357">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="0b7f8-357">
      - Modules</span></span></td>
    <td> <span data-ttu-id="0b7f8-358">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-358">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b7f8-359">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-359">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b7f8-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b7f8-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b7f8-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0b7f8-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="0b7f8-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="0b7f8-365">不可用</span><span class="sxs-lookup"><span data-stu-id="0b7f8-365">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-366">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-366">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-367">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="0b7f8-367">- Mail Read</span></span><br><span data-ttu-id="0b7f8-368">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="0b7f8-368">
      - Mail Compose</span></span><br><span data-ttu-id="0b7f8-369">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-369">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="0b7f8-370">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="0b7f8-370">
      - Modules</span></span></td>
    <td> <span data-ttu-id="0b7f8-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b7f8-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b7f8-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b7f8-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="0b7f8-375">暂无</span><span class="sxs-lookup"><span data-stu-id="0b7f8-375">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-376">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-376">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-377">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="0b7f8-377">- Mail Read</span></span><br><span data-ttu-id="0b7f8-378">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="0b7f8-378">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="0b7f8-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b7f8-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b7f8-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="0b7f8-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="0b7f8-383">暂无</span><span class="sxs-lookup"><span data-stu-id="0b7f8-383">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-384">Office 365 for iOS</span><span class="sxs-lookup"><span data-stu-id="0b7f8-384">Office 365 for iOS</span></span></td>
    <td> <span data-ttu-id="0b7f8-385">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="0b7f8-385">- Mail Read</span></span><br><span data-ttu-id="0b7f8-386">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-386">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-387">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-387">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b7f8-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b7f8-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b7f8-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b7f8-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="0b7f8-392">不可用</span><span class="sxs-lookup"><span data-stu-id="0b7f8-392">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-393">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="0b7f8-393">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="0b7f8-394">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="0b7f8-394">- Mail Read</span></span><br><span data-ttu-id="0b7f8-395">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="0b7f8-395">
      - Mail Compose</span></span><br><span data-ttu-id="0b7f8-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-397">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-397">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b7f8-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b7f8-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b7f8-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b7f8-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0b7f8-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="0b7f8-403">不可用</span><span class="sxs-lookup"><span data-stu-id="0b7f8-403">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-404">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="0b7f8-404">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="0b7f8-405">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="0b7f8-405">- Mail Read</span></span><br><span data-ttu-id="0b7f8-406">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="0b7f8-406">
      - Mail Compose</span></span><br><span data-ttu-id="0b7f8-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b7f8-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b7f8-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b7f8-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b7f8-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0b7f8-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="0b7f8-414">不可用</span><span class="sxs-lookup"><span data-stu-id="0b7f8-414">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-415">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="0b7f8-415">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="0b7f8-416">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="0b7f8-416">- Mail Read</span></span><br><span data-ttu-id="0b7f8-417">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="0b7f8-417">
      - Mail Compose</span></span><br><span data-ttu-id="0b7f8-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-419">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-419">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b7f8-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b7f8-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b7f8-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b7f8-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0b7f8-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="0b7f8-425">不可用</span><span class="sxs-lookup"><span data-stu-id="0b7f8-425">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-426">Office 365 for Android</span><span class="sxs-lookup"><span data-stu-id="0b7f8-426">Office 365 for Android</span></span></td>
    <td> <span data-ttu-id="0b7f8-427">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="0b7f8-427">- Mail Read</span></span><br><span data-ttu-id="0b7f8-428">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-428">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b7f8-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b7f8-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b7f8-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b7f8-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="0b7f8-434">不可用</span><span class="sxs-lookup"><span data-stu-id="0b7f8-434">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="0b7f8-435">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-435">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="0b7f8-436">Word</span><span class="sxs-lookup"><span data-stu-id="0b7f8-436">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="0b7f8-437">平台</span><span class="sxs-lookup"><span data-stu-id="0b7f8-437">Platform</span></span></th>
    <th><span data-ttu-id="0b7f8-438">扩展点</span><span class="sxs-lookup"><span data-stu-id="0b7f8-438">Extension points</span></span></th>
    <th><span data-ttu-id="0b7f8-439">API 要求集</span><span class="sxs-lookup"><span data-stu-id="0b7f8-439">API requirement sets</span></span></th>
    <th><span data-ttu-id="0b7f8-440"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-440"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-441">Office Online</span><span class="sxs-lookup"><span data-stu-id="0b7f8-441">Office Online</span></span></td>
    <td> <span data-ttu-id="0b7f8-442">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-442">- TaskPane</span></span><br><span data-ttu-id="0b7f8-443">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-443">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-444">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-444">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-445">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-445">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0b7f8-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0b7f8-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-448">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-448">- BindingEvents</span></span><br><span data-ttu-id="0b7f8-449">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b7f8-449">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b7f8-450">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-450">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-451">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-451">
         - File</span></span><br><span data-ttu-id="0b7f8-452">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-452">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b7f8-453">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-453">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-454">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-454">
         - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-455">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-455">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-456">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-456">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b7f8-457">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-457">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-458">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-458">
         - Selection</span></span><br><span data-ttu-id="0b7f8-459">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-459">
         - Settings</span></span><br><span data-ttu-id="0b7f8-460">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-460">
         - TableBindings</span></span><br><span data-ttu-id="0b7f8-461">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-461">
         - TableCoercion</span></span><br><span data-ttu-id="0b7f8-462">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-462">
         - TextBindings</span></span><br><span data-ttu-id="0b7f8-463">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-463">
         - TextCoercion</span></span><br><span data-ttu-id="0b7f8-464">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-464">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-465">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-465">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-466">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-466">- TaskPane</span></span><br><span data-ttu-id="0b7f8-467">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-467">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-468">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-468">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-469">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-469">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0b7f8-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0b7f8-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-472">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-472">- BindingEvents</span></span><br><span data-ttu-id="0b7f8-473">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-473">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-474">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b7f8-474">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b7f8-475">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-475">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-476">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-476">
         - File</span></span><br><span data-ttu-id="0b7f8-477">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-477">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b7f8-478">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-478">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-479">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-479">
         - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-480">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-480">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-481">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-481">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b7f8-482">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-482">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-483">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-483">
         - Selection</span></span><br><span data-ttu-id="0b7f8-484">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-484">
         - Settings</span></span><br><span data-ttu-id="0b7f8-485">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-485">
         - TableBindings</span></span><br><span data-ttu-id="0b7f8-486">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-486">
         - TableCoercion</span></span><br><span data-ttu-id="0b7f8-487">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-487">
         - TextBindings</span></span><br><span data-ttu-id="0b7f8-488">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-488">
         - TextCoercion</span></span><br><span data-ttu-id="0b7f8-489">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-489">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-490">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-490">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-491">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-491">- TaskPane</span></span><br><span data-ttu-id="0b7f8-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-493">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-493">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-494">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-494">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0b7f8-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0b7f8-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-497">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-497">- BindingEvents</span></span><br><span data-ttu-id="0b7f8-498">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-498">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-499">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b7f8-499">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b7f8-500">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-500">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-501">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-501">
         - File</span></span><br><span data-ttu-id="0b7f8-502">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-502">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b7f8-503">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-503">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-504">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-504">
         - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-505">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-505">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-506">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-506">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b7f8-507">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-507">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-508">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-508">
         - Selection</span></span><br><span data-ttu-id="0b7f8-509">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-509">
         - Settings</span></span><br><span data-ttu-id="0b7f8-510">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-510">
         - TableBindings</span></span><br><span data-ttu-id="0b7f8-511">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-511">
         - TableCoercion</span></span><br><span data-ttu-id="0b7f8-512">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-512">
         - TextBindings</span></span><br><span data-ttu-id="0b7f8-513">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-513">
         - TextCoercion</span></span><br><span data-ttu-id="0b7f8-514">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-514">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-515">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-515">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-516">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-516">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0b7f8-517">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-517">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-518">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-518">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="0b7f8-519">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-519">- BindingEvents</span></span><br><span data-ttu-id="0b7f8-520">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-520">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-521">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b7f8-521">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b7f8-522">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-522">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-523">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-523">
         - File</span></span><br><span data-ttu-id="0b7f8-524">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-524">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b7f8-525">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-525">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-526">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-526">
         - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-527">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-527">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-528">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-528">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b7f8-529">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-529">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-530">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-530">
         - Selection</span></span><br><span data-ttu-id="0b7f8-531">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-531">
         - Settings</span></span><br><span data-ttu-id="0b7f8-532">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-532">
         - TableBindings</span></span><br><span data-ttu-id="0b7f8-533">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-533">
         - TableCoercion</span></span><br><span data-ttu-id="0b7f8-534">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-534">
         - TextBindings</span></span><br><span data-ttu-id="0b7f8-535">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-535">
         - TextCoercion</span></span><br><span data-ttu-id="0b7f8-536">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-536">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-537">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-537">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-538">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-538">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0b7f8-539">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-539">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="0b7f8-540">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-540">- BindingEvents</span></span><br><span data-ttu-id="0b7f8-541">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-541">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-542">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b7f8-542">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b7f8-543">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-543">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-544">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-544">
         - File</span></span><br><span data-ttu-id="0b7f8-545">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-545">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b7f8-546">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-546">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-547">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-547">
         - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-548">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-548">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-549">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-549">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b7f8-550">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-550">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-551">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-551">
         - Selection</span></span><br><span data-ttu-id="0b7f8-552">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-552">
         - Settings</span></span><br><span data-ttu-id="0b7f8-553">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-553">
         - TableBindings</span></span><br><span data-ttu-id="0b7f8-554">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-554">
         - TableCoercion</span></span><br><span data-ttu-id="0b7f8-555">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-555">
         - TextBindings</span></span><br><span data-ttu-id="0b7f8-556">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-556">
         - TextCoercion</span></span><br><span data-ttu-id="0b7f8-557">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-557">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-558">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="0b7f8-558">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="0b7f8-559">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-559">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0b7f8-560">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-560">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-561">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-561">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0b7f8-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0b7f8-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="0b7f8-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="0b7f8-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-564">- BindingEvents</span></span><br><span data-ttu-id="0b7f8-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-565">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b7f8-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b7f8-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-567">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-568">
         - File</span></span><br><span data-ttu-id="0b7f8-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b7f8-570">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-570">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-571">
         - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b7f8-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-574">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-575">
         - Selection</span></span><br><span data-ttu-id="0b7f8-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-576">
         - Settings</span></span><br><span data-ttu-id="0b7f8-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-577">
         - TableBindings</span></span><br><span data-ttu-id="0b7f8-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-578">
         - TableCoercion</span></span><br><span data-ttu-id="0b7f8-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-579">
         - TextBindings</span></span><br><span data-ttu-id="0b7f8-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-580">
         - TextCoercion</span></span><br><span data-ttu-id="0b7f8-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-582">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="0b7f8-582">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="0b7f8-583">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-583">- TaskPane</span></span><br><span data-ttu-id="0b7f8-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0b7f8-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0b7f8-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="0b7f8-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="0b7f8-589">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-589">- BindingEvents</span></span><br><span data-ttu-id="0b7f8-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-590">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-591">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b7f8-591">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b7f8-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-592">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-593">
         - File</span></span><br><span data-ttu-id="0b7f8-594">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-594">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b7f8-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-595">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-596">
         - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b7f8-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-599">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-600">
         - Selection</span></span><br><span data-ttu-id="0b7f8-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-601">
         - Settings</span></span><br><span data-ttu-id="0b7f8-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-602">
         - TableBindings</span></span><br><span data-ttu-id="0b7f8-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-603">
         - TableCoercion</span></span><br><span data-ttu-id="0b7f8-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-604">
         - TextBindings</span></span><br><span data-ttu-id="0b7f8-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-605">
         - TextCoercion</span></span><br><span data-ttu-id="0b7f8-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-607">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="0b7f8-607">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="0b7f8-608">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-608">- TaskPane</span></span><br><span data-ttu-id="0b7f8-609">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-609">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-611">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-611">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0b7f8-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0b7f8-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="0b7f8-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="0b7f8-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-614">- BindingEvents</span></span><br><span data-ttu-id="0b7f8-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-615">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b7f8-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b7f8-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-617">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-618">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-618">
         - File</span></span><br><span data-ttu-id="0b7f8-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b7f8-620">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-620">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-621">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-621">
         - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-622">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-622">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-623">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-623">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b7f8-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-624">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-625">
         - Selection</span></span><br><span data-ttu-id="0b7f8-626">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-626">
         - Settings</span></span><br><span data-ttu-id="0b7f8-627">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-627">
         - TableBindings</span></span><br><span data-ttu-id="0b7f8-628">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-628">
         - TableCoercion</span></span><br><span data-ttu-id="0b7f8-629">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-629">
         - TextBindings</span></span><br><span data-ttu-id="0b7f8-630">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-630">
         - TextCoercion</span></span><br><span data-ttu-id="0b7f8-631">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-631">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-632">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="0b7f8-632">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="0b7f8-633">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0b7f8-634">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-634">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-635">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-635">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="0b7f8-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-636">- BindingEvents</span></span><br><span data-ttu-id="0b7f8-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-637">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b7f8-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b7f8-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-639">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-640">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-640">
         - File</span></span><br><span data-ttu-id="0b7f8-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b7f8-642">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-642">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-643">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-643">
         - MatrixBindings</span></span><br><span data-ttu-id="0b7f8-644">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-644">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b7f8-645">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-645">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b7f8-646">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-646">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-647">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-647">
         - Selection</span></span><br><span data-ttu-id="0b7f8-648">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-648">
         - Settings</span></span><br><span data-ttu-id="0b7f8-649">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-649">
         - TableBindings</span></span><br><span data-ttu-id="0b7f8-650">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-650">
         - TableCoercion</span></span><br><span data-ttu-id="0b7f8-651">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-651">
         - TextBindings</span></span><br><span data-ttu-id="0b7f8-652">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-652">
         - TextCoercion</span></span><br><span data-ttu-id="0b7f8-653">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-653">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="0b7f8-654">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-654">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="0b7f8-655">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="0b7f8-655">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="0b7f8-656">平台</span><span class="sxs-lookup"><span data-stu-id="0b7f8-656">Platform</span></span></th>
    <th><span data-ttu-id="0b7f8-657">扩展点</span><span class="sxs-lookup"><span data-stu-id="0b7f8-657">Extension points</span></span></th>
    <th><span data-ttu-id="0b7f8-658">API 要求集</span><span class="sxs-lookup"><span data-stu-id="0b7f8-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="0b7f8-659"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-659"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="0b7f8-660">Office Online</span></span></td>
    <td> <span data-ttu-id="0b7f8-661">- 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-661">- Content</span></span><br><span data-ttu-id="0b7f8-662">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-662">
         - TaskPane</span></span><br><span data-ttu-id="0b7f8-663">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-663">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-664">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-664">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-665">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b7f8-665">- ActiveView</span></span><br><span data-ttu-id="0b7f8-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-666">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-667">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-667">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-668">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-668">
         - File</span></span><br><span data-ttu-id="0b7f8-669">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-669">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-670">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-671">
         - Selection</span></span><br><span data-ttu-id="0b7f8-672">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-672">
         - Settings</span></span><br><span data-ttu-id="0b7f8-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-673">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-674">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-674">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-675">- 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-675">- Content</span></span><br><span data-ttu-id="0b7f8-676">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-676">
         - TaskPane</span></span><br><span data-ttu-id="0b7f8-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-678">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-678">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-679">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b7f8-679">- ActiveView</span></span><br><span data-ttu-id="0b7f8-680">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-680">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-681">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-681">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-682">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-682">
         - File</span></span><br><span data-ttu-id="0b7f8-683">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-683">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-684">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-684">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-685">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-685">
         - Selection</span></span><br><span data-ttu-id="0b7f8-686">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-686">
         - Settings</span></span><br><span data-ttu-id="0b7f8-687">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-687">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-688">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-688">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-689">- 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-689">- Content</span></span><br><span data-ttu-id="0b7f8-690">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-690">
         - TaskPane</span></span><br><span data-ttu-id="0b7f8-691">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-691">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-692">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-692">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-693">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b7f8-693">- ActiveView</span></span><br><span data-ttu-id="0b7f8-694">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-694">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-695">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-695">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-696">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-696">
         - File</span></span><br><span data-ttu-id="0b7f8-697">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-697">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-698">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-698">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-699">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-699">
         - Selection</span></span><br><span data-ttu-id="0b7f8-700">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-700">
         - Settings</span></span><br><span data-ttu-id="0b7f8-701">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-701">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-702">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-702">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-703">- 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-703">- Content</span></span><br><span data-ttu-id="0b7f8-704">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-704">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="0b7f8-705">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-705">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="0b7f8-706">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b7f8-706">- ActiveView</span></span><br><span data-ttu-id="0b7f8-707">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-707">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-708">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-708">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-709">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-709">
         - File</span></span><br><span data-ttu-id="0b7f8-710">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-710">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-711">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-711">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-712">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-712">
         - Selection</span></span><br><span data-ttu-id="0b7f8-713">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-713">
         - Settings</span></span><br><span data-ttu-id="0b7f8-714">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-714">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-715">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-715">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-716">- 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-716">- Content</span></span><br><span data-ttu-id="0b7f8-717">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-717">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="0b7f8-718">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-718">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="0b7f8-719">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b7f8-719">- ActiveView</span></span><br><span data-ttu-id="0b7f8-720">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-720">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-721">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-721">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-722">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-722">
         - File</span></span><br><span data-ttu-id="0b7f8-723">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-723">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-724">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-724">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-725">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-725">
         - Selection</span></span><br><span data-ttu-id="0b7f8-726">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-726">
         - Settings</span></span><br><span data-ttu-id="0b7f8-727">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-727">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-728">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="0b7f8-728">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="0b7f8-729">- 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-729">- Content</span></span><br><span data-ttu-id="0b7f8-730">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-730">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="0b7f8-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="0b7f8-732">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b7f8-732">- ActiveView</span></span><br><span data-ttu-id="0b7f8-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-733">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-734">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-734">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-735">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-735">
         - File</span></span><br><span data-ttu-id="0b7f8-736">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-736">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-737">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-737">
         - Selection</span></span><br><span data-ttu-id="0b7f8-738">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-738">
         - Settings</span></span><br><span data-ttu-id="0b7f8-739">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-739">
         - TextCoercion</span></span><br><span data-ttu-id="0b7f8-740">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-740">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-741">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="0b7f8-741">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="0b7f8-742">- 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-742">- Content</span></span><br><span data-ttu-id="0b7f8-743">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-743">
         - TaskPane</span></span><br><span data-ttu-id="0b7f8-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-745">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-745">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-746">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b7f8-746">- ActiveView</span></span><br><span data-ttu-id="0b7f8-747">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-747">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-748">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-748">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-749">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-749">
         - File</span></span><br><span data-ttu-id="0b7f8-750">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-750">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-751">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-751">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-752">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-752">
         - Selection</span></span><br><span data-ttu-id="0b7f8-753">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-753">
         - Settings</span></span><br><span data-ttu-id="0b7f8-754">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-754">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-755">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="0b7f8-755">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="0b7f8-756">- 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-756">- Content</span></span><br><span data-ttu-id="0b7f8-757">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-757">
         - TaskPane</span></span><br><span data-ttu-id="0b7f8-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-759">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-759">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-760">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b7f8-760">- ActiveView</span></span><br><span data-ttu-id="0b7f8-761">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-761">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-762">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-762">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-763">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-763">
         - File</span></span><br><span data-ttu-id="0b7f8-764">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-764">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-765">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-765">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-766">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-766">
         - Selection</span></span><br><span data-ttu-id="0b7f8-767">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-767">
         - Settings</span></span><br><span data-ttu-id="0b7f8-768">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-768">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-769">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="0b7f8-769">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="0b7f8-770">- 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-770">- Content</span></span><br><span data-ttu-id="0b7f8-771">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-771">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="0b7f8-772">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-772">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="0b7f8-773">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b7f8-773">- ActiveView</span></span><br><span data-ttu-id="0b7f8-774">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-774">
         - CompressedFile</span></span><br><span data-ttu-id="0b7f8-775">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-775">
         - DocumentEvents</span></span><br><span data-ttu-id="0b7f8-776">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b7f8-776">
         - File</span></span><br><span data-ttu-id="0b7f8-777">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-777">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-778">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b7f8-778">
         - PdfFile</span></span><br><span data-ttu-id="0b7f8-779">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-779">
         - Selection</span></span><br><span data-ttu-id="0b7f8-780">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-780">
         - Settings</span></span><br><span data-ttu-id="0b7f8-781">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-781">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="0b7f8-782">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="0b7f8-782">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="0b7f8-783">OneNote</span><span class="sxs-lookup"><span data-stu-id="0b7f8-783">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="0b7f8-784">平台</span><span class="sxs-lookup"><span data-stu-id="0b7f8-784">Platform</span></span></th>
    <th><span data-ttu-id="0b7f8-785">扩展点</span><span class="sxs-lookup"><span data-stu-id="0b7f8-785">Extension points</span></span></th>
    <th><span data-ttu-id="0b7f8-786">API 要求集</span><span class="sxs-lookup"><span data-stu-id="0b7f8-786">API requirement sets</span></span></th>
    <th><span data-ttu-id="0b7f8-787"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-787"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-788">Office Online</span><span class="sxs-lookup"><span data-stu-id="0b7f8-788">Office Online</span></span></td>
    <td> <span data-ttu-id="0b7f8-789">- 内容</span><span class="sxs-lookup"><span data-stu-id="0b7f8-789">- Content</span></span><br><span data-ttu-id="0b7f8-790">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-790">
         - TaskPane</span></span><br><span data-ttu-id="0b7f8-791">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-791">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-792">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-792">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="0b7f8-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-794">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b7f8-794">- DocumentEvents</span></span><br><span data-ttu-id="0b7f8-795">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-795">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b7f8-796">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-796">
         - ImageCoercion</span></span><br><span data-ttu-id="0b7f8-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b7f8-797">
         - Settings</span></span><br><span data-ttu-id="0b7f8-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-798">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="0b7f8-799">项目</span><span class="sxs-lookup"><span data-stu-id="0b7f8-799">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="0b7f8-800">平台</span><span class="sxs-lookup"><span data-stu-id="0b7f8-800">Platform</span></span></th>
    <th><span data-ttu-id="0b7f8-801">扩展点</span><span class="sxs-lookup"><span data-stu-id="0b7f8-801">Extension points</span></span></th>
    <th><span data-ttu-id="0b7f8-802">API 要求集</span><span class="sxs-lookup"><span data-stu-id="0b7f8-802">API requirement sets</span></span></th>
    <th><span data-ttu-id="0b7f8-803"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-803"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-804">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-804">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-805">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-805">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0b7f8-806">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-806">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-807">- Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-807">- Selection</span></span><br><span data-ttu-id="0b7f8-808">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-808">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-809">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-809">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-810">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-810">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0b7f8-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-812">- Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-812">- Selection</span></span><br><span data-ttu-id="0b7f8-813">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-813">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b7f8-814">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="0b7f8-814">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="0b7f8-815">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="0b7f8-815">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0b7f8-816">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b7f8-816">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b7f8-817">- Selection</span><span class="sxs-lookup"><span data-stu-id="0b7f8-817">- Selection</span></span><br><span data-ttu-id="0b7f8-818">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b7f8-818">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="0b7f8-819">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0b7f8-819">See also</span></span>

- [<span data-ttu-id="0b7f8-820">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="0b7f8-820">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="0b7f8-821">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="0b7f8-821">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="0b7f8-822">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="0b7f8-822">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="0b7f8-823">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="0b7f8-823">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="0b7f8-824">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="0b7f8-824">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="0b7f8-825">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="0b7f8-825">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="0b7f8-826">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="0b7f8-826">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="0b7f8-827">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="0b7f8-827">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
---
title: Office 外接程序主机和平台可用性
description: Excel、Word、Outlook、PowerPoint、OneNote 和项目支持的要求集。
ms.date: 05/23/2019
localization_priority: Priority
ms.openlocfilehash: 6fb1f0db839910e91d7a5215f8e21f5b33ff2165
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432192"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="207d0-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="207d0-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="207d0-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="207d0-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="207d0-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="207d0-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="207d0-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="207d0-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="207d0-108">Excel</span><span class="sxs-lookup"><span data-stu-id="207d0-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="207d0-109">平台</span><span class="sxs-lookup"><span data-stu-id="207d0-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="207d0-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="207d0-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="207d0-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="207d0-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="207d0-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="207d0-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="207d0-113">Office Online</span></span></td>
    <td> <span data-ttu-id="207d0-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-114">- TaskPane</span></span><br><span data-ttu-id="207d0-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-115">
        - Content</span></span><br><span data-ttu-id="207d0-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="207d0-116">
        - Custom Functions</span></span><br><span data-ttu-id="207d0-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="207d0-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="207d0-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="207d0-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="207d0-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="207d0-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="207d0-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="207d0-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="207d0-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="207d0-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="207d0-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="207d0-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="207d0-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="207d0-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="207d0-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="207d0-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="207d0-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="207d0-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="207d0-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-128">
        - BindingEvents</span></span><br><span data-ttu-id="207d0-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-129">
        - CompressedFile</span></span><br><span data-ttu-id="207d0-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-130">
        - DocumentEvents</span></span><br><span data-ttu-id="207d0-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="207d0-131">
        - File</span></span><br><span data-ttu-id="207d0-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-132">
        - MatrixBindings</span></span><br><span data-ttu-id="207d0-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="207d0-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-134">
        - Selection</span></span><br><span data-ttu-id="207d0-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-135">
        - Settings</span></span><br><span data-ttu-id="207d0-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-136">
        - TableBindings</span></span><br><span data-ttu-id="207d0-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-137">
        - TableCoercion</span></span><br><span data-ttu-id="207d0-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-138">
        - TextBindings</span></span><br><span data-ttu-id="207d0-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-140">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="207d0-140">Office on Windows</span></span><br><span data-ttu-id="207d0-141">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-141">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="207d0-142">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-142">- TaskPane</span></span><br><span data-ttu-id="207d0-143">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-143">
        - Content</span></span><br><span data-ttu-id="207d0-144">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="207d0-144">
        - Custom Functions</span></span><br><span data-ttu-id="207d0-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="207d0-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="207d0-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="207d0-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="207d0-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="207d0-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="207d0-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="207d0-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="207d0-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="207d0-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="207d0-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="207d0-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="207d0-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="207d0-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="207d0-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="207d0-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="207d0-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="207d0-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="207d0-156">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-156">
        - BindingEvents</span></span><br><span data-ttu-id="207d0-157">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-157">
        - CompressedFile</span></span><br><span data-ttu-id="207d0-158">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-158">
        - DocumentEvents</span></span><br><span data-ttu-id="207d0-159">
        - File</span><span class="sxs-lookup"><span data-stu-id="207d0-159">
        - File</span></span><br><span data-ttu-id="207d0-160">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-160">
        - MatrixBindings</span></span><br><span data-ttu-id="207d0-161">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-161">
        - MatrixCoercion</span></span><br><span data-ttu-id="207d0-162">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-162">
        - Selection</span></span><br><span data-ttu-id="207d0-163">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-163">
        - Settings</span></span><br><span data-ttu-id="207d0-164">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-164">
        - TableBindings</span></span><br><span data-ttu-id="207d0-165">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-165">
        - TableCoercion</span></span><br><span data-ttu-id="207d0-166">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-166">
        - TextBindings</span></span><br><span data-ttu-id="207d0-167">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-167">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-168">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="207d0-168">Office 2019 on Windows</span></span><br><span data-ttu-id="207d0-169">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-169">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="207d0-170">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-170">- TaskPane</span></span><br><span data-ttu-id="207d0-171">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-171">
        - Content</span></span><br><span data-ttu-id="207d0-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="207d0-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="207d0-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="207d0-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="207d0-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="207d0-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="207d0-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="207d0-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="207d0-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="207d0-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="207d0-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="207d0-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="207d0-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="207d0-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="207d0-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="207d0-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-182">- BindingEvents</span></span><br><span data-ttu-id="207d0-183">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-183">
        - CompressedFile</span></span><br><span data-ttu-id="207d0-184">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-184">
        - DocumentEvents</span></span><br><span data-ttu-id="207d0-185">
        - File</span><span class="sxs-lookup"><span data-stu-id="207d0-185">
        - File</span></span><br><span data-ttu-id="207d0-186">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-186">
        - ImageCoercion</span></span><br><span data-ttu-id="207d0-187">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-187">
        - MatrixBindings</span></span><br><span data-ttu-id="207d0-188">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-188">
        - MatrixCoercion</span></span><br><span data-ttu-id="207d0-189">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-189">
        - Selection</span></span><br><span data-ttu-id="207d0-190">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-190">
        - Settings</span></span><br><span data-ttu-id="207d0-191">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-191">
        - TableBindings</span></span><br><span data-ttu-id="207d0-192">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-192">
        - TableCoercion</span></span><br><span data-ttu-id="207d0-193">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-193">
        - TextBindings</span></span><br><span data-ttu-id="207d0-194">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-194">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-195">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="207d0-195">Office 2016 on Windows</span></span><br><span data-ttu-id="207d0-196">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-196">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="207d0-197">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-197">- TaskPane</span></span><br><span data-ttu-id="207d0-198">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-198">
        - Content</span></span></td>
    <td><span data-ttu-id="207d0-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="207d0-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="207d0-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="207d0-201">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-201">- BindingEvents</span></span><br><span data-ttu-id="207d0-202">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-202">
        - CompressedFile</span></span><br><span data-ttu-id="207d0-203">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-203">
        - DocumentEvents</span></span><br><span data-ttu-id="207d0-204">
        - File</span><span class="sxs-lookup"><span data-stu-id="207d0-204">
        - File</span></span><br><span data-ttu-id="207d0-205">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-205">
        - ImageCoercion</span></span><br><span data-ttu-id="207d0-206">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-206">
        - MatrixBindings</span></span><br><span data-ttu-id="207d0-207">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-207">
        - MatrixCoercion</span></span><br><span data-ttu-id="207d0-208">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-208">
        - Selection</span></span><br><span data-ttu-id="207d0-209">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-209">
        - Settings</span></span><br><span data-ttu-id="207d0-210">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-210">
        - TableBindings</span></span><br><span data-ttu-id="207d0-211">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-211">
        - TableCoercion</span></span><br><span data-ttu-id="207d0-212">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-212">
        - TextBindings</span></span><br><span data-ttu-id="207d0-213">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-213">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-214">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="207d0-214">Office 2013 on Windows</span></span><br><span data-ttu-id="207d0-215">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-215">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="207d0-216">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-216">
        - TaskPane</span></span><br><span data-ttu-id="207d0-217">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-217">
        - Content</span></span></td>
    <td>  <span data-ttu-id="207d0-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="207d0-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="207d0-219">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-219">
        - BindingEvents</span></span><br><span data-ttu-id="207d0-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-220">
        - CompressedFile</span></span><br><span data-ttu-id="207d0-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-221">
        - DocumentEvents</span></span><br><span data-ttu-id="207d0-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="207d0-222">
        - File</span></span><br><span data-ttu-id="207d0-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-223">
        - ImageCoercion</span></span><br><span data-ttu-id="207d0-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-224">
        - MatrixBindings</span></span><br><span data-ttu-id="207d0-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="207d0-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-226">
        - Selection</span></span><br><span data-ttu-id="207d0-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-227">
        - Settings</span></span><br><span data-ttu-id="207d0-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-228">
        - TableBindings</span></span><br><span data-ttu-id="207d0-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-229">
        - TableCoercion</span></span><br><span data-ttu-id="207d0-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-230">
        - TextBindings</span></span><br><span data-ttu-id="207d0-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-232">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="207d0-232">Office for iPad</span></span><br><span data-ttu-id="207d0-233">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-233">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="207d0-234">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-234">- TaskPane</span></span><br><span data-ttu-id="207d0-235">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-235">
        - Content</span></span><br><span data-ttu-id="207d0-236">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="207d0-236">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="207d0-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="207d0-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="207d0-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="207d0-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="207d0-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="207d0-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="207d0-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="207d0-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="207d0-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="207d0-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="207d0-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="207d0-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="207d0-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="207d0-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="207d0-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="207d0-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="207d0-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-247">- BindingEvents</span></span><br><span data-ttu-id="207d0-248">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-248">
        - DocumentEvents</span></span><br><span data-ttu-id="207d0-249">
        - File</span><span class="sxs-lookup"><span data-stu-id="207d0-249">
        - File</span></span><br><span data-ttu-id="207d0-250">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-250">
        - ImageCoercion</span></span><br><span data-ttu-id="207d0-251">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-251">
        - MatrixBindings</span></span><br><span data-ttu-id="207d0-252">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-252">
        - MatrixCoercion</span></span><br><span data-ttu-id="207d0-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-253">
        - Selection</span></span><br><span data-ttu-id="207d0-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-254">
        - Settings</span></span><br><span data-ttu-id="207d0-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-255">
        - TableBindings</span></span><br><span data-ttu-id="207d0-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-256">
        - TableCoercion</span></span><br><span data-ttu-id="207d0-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-257">
        - TextBindings</span></span><br><span data-ttu-id="207d0-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-259">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="207d0-259">Office for Mac</span></span><br><span data-ttu-id="207d0-260">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-260">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="207d0-261">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-261">- TaskPane</span></span><br><span data-ttu-id="207d0-262">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-262">
        - Content</span></span><br><span data-ttu-id="207d0-263">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="207d0-263">
        - Custom Functions</span></span><br><span data-ttu-id="207d0-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="207d0-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="207d0-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="207d0-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="207d0-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="207d0-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="207d0-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="207d0-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="207d0-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="207d0-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="207d0-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="207d0-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="207d0-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="207d0-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="207d0-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="207d0-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="207d0-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="207d0-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-275">- BindingEvents</span></span><br><span data-ttu-id="207d0-276">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-276">
        - CompressedFile</span></span><br><span data-ttu-id="207d0-277">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-277">
        - DocumentEvents</span></span><br><span data-ttu-id="207d0-278">
        - File</span><span class="sxs-lookup"><span data-stu-id="207d0-278">
        - File</span></span><br><span data-ttu-id="207d0-279">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-279">
        - ImageCoercion</span></span><br><span data-ttu-id="207d0-280">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-280">
        - MatrixBindings</span></span><br><span data-ttu-id="207d0-281">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-281">
        - MatrixCoercion</span></span><br><span data-ttu-id="207d0-282">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-282">
        - PdfFile</span></span><br><span data-ttu-id="207d0-283">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-283">
        - Selection</span></span><br><span data-ttu-id="207d0-284">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-284">
        - Settings</span></span><br><span data-ttu-id="207d0-285">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-285">
        - TableBindings</span></span><br><span data-ttu-id="207d0-286">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-286">
        - TableCoercion</span></span><br><span data-ttu-id="207d0-287">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-287">
        - TextBindings</span></span><br><span data-ttu-id="207d0-288">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-288">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-289">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="207d0-289">Office 2019 for Mac</span></span><br><span data-ttu-id="207d0-290">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-290">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="207d0-291">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-291">- TaskPane</span></span><br><span data-ttu-id="207d0-292">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-292">
        - Content</span></span><br><span data-ttu-id="207d0-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="207d0-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="207d0-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="207d0-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="207d0-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="207d0-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="207d0-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="207d0-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="207d0-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="207d0-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="207d0-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="207d0-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="207d0-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="207d0-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="207d0-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="207d0-303">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-303">- BindingEvents</span></span><br><span data-ttu-id="207d0-304">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-304">
        - CompressedFile</span></span><br><span data-ttu-id="207d0-305">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-305">
        - DocumentEvents</span></span><br><span data-ttu-id="207d0-306">
        - File</span><span class="sxs-lookup"><span data-stu-id="207d0-306">
        - File</span></span><br><span data-ttu-id="207d0-307">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-307">
        - ImageCoercion</span></span><br><span data-ttu-id="207d0-308">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-308">
        - MatrixBindings</span></span><br><span data-ttu-id="207d0-309">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-309">
        - MatrixCoercion</span></span><br><span data-ttu-id="207d0-310">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-310">
        - PdfFile</span></span><br><span data-ttu-id="207d0-311">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-311">
        - Selection</span></span><br><span data-ttu-id="207d0-312">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-312">
        - Settings</span></span><br><span data-ttu-id="207d0-313">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-313">
        - TableBindings</span></span><br><span data-ttu-id="207d0-314">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-314">
        - TableCoercion</span></span><br><span data-ttu-id="207d0-315">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-315">
        - TextBindings</span></span><br><span data-ttu-id="207d0-316">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-316">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-317">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="207d0-317">Office 2016 for Mac</span></span><br><span data-ttu-id="207d0-318">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-318">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="207d0-319">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-319">- TaskPane</span></span><br><span data-ttu-id="207d0-320">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-320">
        - Content</span></span></td>
    <td><span data-ttu-id="207d0-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="207d0-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="207d0-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="207d0-323">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-323">- BindingEvents</span></span><br><span data-ttu-id="207d0-324">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-324">
        - CompressedFile</span></span><br><span data-ttu-id="207d0-325">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-325">
        - DocumentEvents</span></span><br><span data-ttu-id="207d0-326">
        - File</span><span class="sxs-lookup"><span data-stu-id="207d0-326">
        - File</span></span><br><span data-ttu-id="207d0-327">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-327">
        - ImageCoercion</span></span><br><span data-ttu-id="207d0-328">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-328">
        - MatrixBindings</span></span><br><span data-ttu-id="207d0-329">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-329">
        - MatrixCoercion</span></span><br><span data-ttu-id="207d0-330">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-330">
        - PdfFile</span></span><br><span data-ttu-id="207d0-331">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-331">
        - Selection</span></span><br><span data-ttu-id="207d0-332">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-332">
        - Settings</span></span><br><span data-ttu-id="207d0-333">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-333">
        - TableBindings</span></span><br><span data-ttu-id="207d0-334">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-334">
        - TableCoercion</span></span><br><span data-ttu-id="207d0-335">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-335">
        - TextBindings</span></span><br><span data-ttu-id="207d0-336">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-336">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="207d0-337">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="207d0-337">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="207d0-338">自定义函数</span><span class="sxs-lookup"><span data-stu-id="207d0-338">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="207d0-339">平台</span><span class="sxs-lookup"><span data-stu-id="207d0-339">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="207d0-340">扩展点</span><span class="sxs-lookup"><span data-stu-id="207d0-340">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="207d0-341">API 要求集</span><span class="sxs-lookup"><span data-stu-id="207d0-341">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="207d0-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="207d0-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-343">Office Online</span><span class="sxs-lookup"><span data-stu-id="207d0-343">Office Online</span></span></td>
    <td><span data-ttu-id="207d0-344">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="207d0-344">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="207d0-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-346">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="207d0-346">Office on Windows</span></span><br><span data-ttu-id="207d0-347">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-347">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="207d0-348">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="207d0-348">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="207d0-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-350">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="207d0-350">Office for iPad</span></span><br><span data-ttu-id="207d0-351">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-351">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="207d0-352">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="207d0-352">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="207d0-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-354">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="207d0-354">Office for Mac</span></span><br><span data-ttu-id="207d0-355">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-355">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="207d0-356">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="207d0-356">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="207d0-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="207d0-358">Outlook</span><span class="sxs-lookup"><span data-stu-id="207d0-358">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="207d0-359">平台</span><span class="sxs-lookup"><span data-stu-id="207d0-359">Platform</span></span></th>
    <th><span data-ttu-id="207d0-360">扩展点</span><span class="sxs-lookup"><span data-stu-id="207d0-360">Extension points</span></span></th>
    <th><span data-ttu-id="207d0-361">API 要求集</span><span class="sxs-lookup"><span data-stu-id="207d0-361">API requirement sets</span></span></th>
    <th><span data-ttu-id="207d0-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="207d0-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-363">Office Online</span><span class="sxs-lookup"><span data-stu-id="207d0-363">Office Online</span></span></td>
    <td> <span data-ttu-id="207d0-364">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="207d0-364">- Mail Read</span></span><br><span data-ttu-id="207d0-365">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="207d0-365">
      - Mail Compose</span></span><br><span data-ttu-id="207d0-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="207d0-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="207d0-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="207d0-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="207d0-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="207d0-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="207d0-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="207d0-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="207d0-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="207d0-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="207d0-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="207d0-374">不可用</span><span class="sxs-lookup"><span data-stu-id="207d0-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-375">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="207d0-375">Office on Windows</span></span><br><span data-ttu-id="207d0-376">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-376">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="207d0-377">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="207d0-377">- Mail Read</span></span><br><span data-ttu-id="207d0-378">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="207d0-378">
      - Mail Compose</span></span><br><span data-ttu-id="207d0-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="207d0-380">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="207d0-380">
      - Modules</span></span></td>
    <td> <span data-ttu-id="207d0-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="207d0-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="207d0-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="207d0-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="207d0-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="207d0-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="207d0-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="207d0-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="207d0-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="207d0-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="207d0-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="207d0-388">不可用</span><span class="sxs-lookup"><span data-stu-id="207d0-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-389">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="207d0-389">Office 2019 on Windows</span></span><br><span data-ttu-id="207d0-390">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-390">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-391">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="207d0-391">- Mail Read</span></span><br><span data-ttu-id="207d0-392">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="207d0-392">
      - Mail Compose</span></span><br><span data-ttu-id="207d0-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="207d0-394">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="207d0-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="207d0-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="207d0-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="207d0-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="207d0-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="207d0-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="207d0-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="207d0-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="207d0-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="207d0-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="207d0-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="207d0-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="207d0-402">不可用</span><span class="sxs-lookup"><span data-stu-id="207d0-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-403">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="207d0-403">Office 2016 on Windows</span></span><br><span data-ttu-id="207d0-404">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-405">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="207d0-405">- Mail Read</span></span><br><span data-ttu-id="207d0-406">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="207d0-406">
      - Mail Compose</span></span><br><span data-ttu-id="207d0-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="207d0-408">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="207d0-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="207d0-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="207d0-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="207d0-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="207d0-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="207d0-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="207d0-413">不可用</span><span class="sxs-lookup"><span data-stu-id="207d0-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-414">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="207d0-414">Office 2013 on Windows</span></span><br><span data-ttu-id="207d0-415">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-416">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="207d0-416">- Mail Read</span></span><br><span data-ttu-id="207d0-417">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="207d0-417">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="207d0-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="207d0-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="207d0-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="207d0-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="207d0-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="207d0-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="207d0-422">不可用</span><span class="sxs-lookup"><span data-stu-id="207d0-422">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-423">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="207d0-423">Office for iOS</span></span><br><span data-ttu-id="207d0-424">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-424">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="207d0-425">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="207d0-425">- Mail Read</span></span><br><span data-ttu-id="207d0-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="207d0-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="207d0-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="207d0-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="207d0-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="207d0-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="207d0-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="207d0-432">不可用</span><span class="sxs-lookup"><span data-stu-id="207d0-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-433">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="207d0-433">Office for Mac</span></span><br><span data-ttu-id="207d0-434">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-434">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="207d0-435">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="207d0-435">- Mail Read</span></span><br><span data-ttu-id="207d0-436">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="207d0-436">
      - Mail Compose</span></span><br><span data-ttu-id="207d0-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="207d0-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="207d0-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="207d0-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="207d0-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="207d0-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="207d0-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="207d0-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="207d0-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="207d0-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="207d0-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="207d0-445">不可用</span><span class="sxs-lookup"><span data-stu-id="207d0-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-446">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="207d0-446">Office 2019 for Mac</span></span><br><span data-ttu-id="207d0-447">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-447">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-448">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="207d0-448">- Mail Read</span></span><br><span data-ttu-id="207d0-449">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="207d0-449">
      - Mail Compose</span></span><br><span data-ttu-id="207d0-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="207d0-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="207d0-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="207d0-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="207d0-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="207d0-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="207d0-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="207d0-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="207d0-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="207d0-457">不可用</span><span class="sxs-lookup"><span data-stu-id="207d0-457">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-458">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="207d0-458">Office 2016 for Mac</span></span><br><span data-ttu-id="207d0-459">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-459">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-460">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="207d0-460">- Mail Read</span></span><br><span data-ttu-id="207d0-461">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="207d0-461">
      - Mail Compose</span></span><br><span data-ttu-id="207d0-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="207d0-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="207d0-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="207d0-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="207d0-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="207d0-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="207d0-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="207d0-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="207d0-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="207d0-469">不可用</span><span class="sxs-lookup"><span data-stu-id="207d0-469">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-470">Office for Android</span><span class="sxs-lookup"><span data-stu-id="207d0-470">Office for Android</span></span><br><span data-ttu-id="207d0-471">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-471">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="207d0-472">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="207d0-472">- Mail Read</span></span><br><span data-ttu-id="207d0-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="207d0-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="207d0-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="207d0-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="207d0-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="207d0-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="207d0-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="207d0-479">不可用</span><span class="sxs-lookup"><span data-stu-id="207d0-479">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="207d0-480">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="207d0-480">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="207d0-481">Word</span><span class="sxs-lookup"><span data-stu-id="207d0-481">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="207d0-482">平台</span><span class="sxs-lookup"><span data-stu-id="207d0-482">Platform</span></span></th>
    <th><span data-ttu-id="207d0-483">扩展点</span><span class="sxs-lookup"><span data-stu-id="207d0-483">Extension points</span></span></th>
    <th><span data-ttu-id="207d0-484">API 要求集</span><span class="sxs-lookup"><span data-stu-id="207d0-484">API requirement sets</span></span></th>
    <th><span data-ttu-id="207d0-485"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="207d0-485"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-486">Office Online</span><span class="sxs-lookup"><span data-stu-id="207d0-486">Office Online</span></span></td>
    <td> <span data-ttu-id="207d0-487">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-487">- TaskPane</span></span><br><span data-ttu-id="207d0-488">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-488">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-489">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-489">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="207d0-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="207d0-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="207d0-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="207d0-493">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-493">- BindingEvents</span></span><br><span data-ttu-id="207d0-494">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="207d0-494">
         - CustomXmlParts</span></span><br><span data-ttu-id="207d0-495">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-495">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-496">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-496">
         - File</span></span><br><span data-ttu-id="207d0-497">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-497">
         - HtmlCoercion</span></span><br><span data-ttu-id="207d0-498">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-498">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-499">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-499">
         - MatrixBindings</span></span><br><span data-ttu-id="207d0-500">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-500">
         - MatrixCoercion</span></span><br><span data-ttu-id="207d0-501">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-501">
         - OoxmlCoercion</span></span><br><span data-ttu-id="207d0-502">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-502">
         - PdfFile</span></span><br><span data-ttu-id="207d0-503">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-503">
         - Selection</span></span><br><span data-ttu-id="207d0-504">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-504">
         - Settings</span></span><br><span data-ttu-id="207d0-505">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-505">
         - TableBindings</span></span><br><span data-ttu-id="207d0-506">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-506">
         - TableCoercion</span></span><br><span data-ttu-id="207d0-507">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-507">
         - TextBindings</span></span><br><span data-ttu-id="207d0-508">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-508">
         - TextCoercion</span></span><br><span data-ttu-id="207d0-509">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="207d0-509">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-510">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="207d0-510">Office on Windows</span></span><br><span data-ttu-id="207d0-511">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-511">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="207d0-512">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-512">- TaskPane</span></span><br><span data-ttu-id="207d0-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-514">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-514">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="207d0-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="207d0-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="207d0-517">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-517">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="207d0-518">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-518">- BindingEvents</span></span><br><span data-ttu-id="207d0-519">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-519">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-520">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="207d0-520">
         - CustomXmlParts</span></span><br><span data-ttu-id="207d0-521">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-521">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-522">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-522">
         - File</span></span><br><span data-ttu-id="207d0-523">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-523">
         - HtmlCoercion</span></span><br><span data-ttu-id="207d0-524">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-524">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-525">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-525">
         - MatrixBindings</span></span><br><span data-ttu-id="207d0-526">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-526">
         - MatrixCoercion</span></span><br><span data-ttu-id="207d0-527">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-527">
         - OoxmlCoercion</span></span><br><span data-ttu-id="207d0-528">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-528">
         - PdfFile</span></span><br><span data-ttu-id="207d0-529">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-529">
         - Selection</span></span><br><span data-ttu-id="207d0-530">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-530">
         - Settings</span></span><br><span data-ttu-id="207d0-531">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-531">
         - TableBindings</span></span><br><span data-ttu-id="207d0-532">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-532">
         - TableCoercion</span></span><br><span data-ttu-id="207d0-533">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-533">
         - TextBindings</span></span><br><span data-ttu-id="207d0-534">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-534">
         - TextCoercion</span></span><br><span data-ttu-id="207d0-535">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="207d0-535">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-536">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="207d0-536">Office 2019 on Windows</span></span><br><span data-ttu-id="207d0-537">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-537">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-538">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-538">- TaskPane</span></span><br><span data-ttu-id="207d0-539">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-539">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-540">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-540">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="207d0-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="207d0-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="207d0-543">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-543">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="207d0-544">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-544">- BindingEvents</span></span><br><span data-ttu-id="207d0-545">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-545">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-546">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="207d0-546">
         - CustomXmlParts</span></span><br><span data-ttu-id="207d0-547">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-547">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-548">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-548">
         - File</span></span><br><span data-ttu-id="207d0-549">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-549">
         - HtmlCoercion</span></span><br><span data-ttu-id="207d0-550">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-550">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-551">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-551">
         - MatrixBindings</span></span><br><span data-ttu-id="207d0-552">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-552">
         - MatrixCoercion</span></span><br><span data-ttu-id="207d0-553">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-553">
         - OoxmlCoercion</span></span><br><span data-ttu-id="207d0-554">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-554">
         - PdfFile</span></span><br><span data-ttu-id="207d0-555">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-555">
         - Selection</span></span><br><span data-ttu-id="207d0-556">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-556">
         - Settings</span></span><br><span data-ttu-id="207d0-557">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-557">
         - TableBindings</span></span><br><span data-ttu-id="207d0-558">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-558">
         - TableCoercion</span></span><br><span data-ttu-id="207d0-559">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-559">
         - TextBindings</span></span><br><span data-ttu-id="207d0-560">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-560">
         - TextCoercion</span></span><br><span data-ttu-id="207d0-561">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="207d0-561">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-562">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="207d0-562">Office 2016 on Windows</span></span><br><span data-ttu-id="207d0-563">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-563">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-564">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-564">- TaskPane</span></span></td>
    <td> <span data-ttu-id="207d0-565">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-565">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="207d0-566">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="207d0-566">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="207d0-567">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-567">- BindingEvents</span></span><br><span data-ttu-id="207d0-568">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-568">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-569">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="207d0-569">
         - CustomXmlParts</span></span><br><span data-ttu-id="207d0-570">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-570">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-571">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-571">
         - File</span></span><br><span data-ttu-id="207d0-572">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-572">
         - HtmlCoercion</span></span><br><span data-ttu-id="207d0-573">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-573">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-574">
         - MatrixBindings</span></span><br><span data-ttu-id="207d0-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="207d0-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="207d0-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-577">
         - PdfFile</span></span><br><span data-ttu-id="207d0-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-578">
         - Selection</span></span><br><span data-ttu-id="207d0-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-579">
         - Settings</span></span><br><span data-ttu-id="207d0-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-580">
         - TableBindings</span></span><br><span data-ttu-id="207d0-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-581">
         - TableCoercion</span></span><br><span data-ttu-id="207d0-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-582">
         - TextBindings</span></span><br><span data-ttu-id="207d0-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-583">
         - TextCoercion</span></span><br><span data-ttu-id="207d0-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="207d0-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-585">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="207d0-585">Office 2013 on Windows</span></span><br><span data-ttu-id="207d0-586">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-587">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="207d0-588">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="207d0-588">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="207d0-589">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-589">- BindingEvents</span></span><br><span data-ttu-id="207d0-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-590">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-591">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="207d0-591">
         - CustomXmlParts</span></span><br><span data-ttu-id="207d0-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-592">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-593">
         - File</span></span><br><span data-ttu-id="207d0-594">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-594">
         - HtmlCoercion</span></span><br><span data-ttu-id="207d0-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-595">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-596">
         - MatrixBindings</span></span><br><span data-ttu-id="207d0-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="207d0-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="207d0-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-599">
         - PdfFile</span></span><br><span data-ttu-id="207d0-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-600">
         - Selection</span></span><br><span data-ttu-id="207d0-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-601">
         - Settings</span></span><br><span data-ttu-id="207d0-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-602">
         - TableBindings</span></span><br><span data-ttu-id="207d0-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-603">
         - TableCoercion</span></span><br><span data-ttu-id="207d0-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-604">
         - TextBindings</span></span><br><span data-ttu-id="207d0-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-605">
         - TextCoercion</span></span><br><span data-ttu-id="207d0-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="207d0-606">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-607">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="207d0-607">Office for iPad</span></span><br><span data-ttu-id="207d0-608">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-608">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="207d0-609">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-609">- TaskPane</span></span></td>
    <td> <span data-ttu-id="207d0-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="207d0-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="207d0-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="207d0-613">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="207d0-613">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="207d0-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-614">- BindingEvents</span></span><br><span data-ttu-id="207d0-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-615">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="207d0-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="207d0-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-617">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-618">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-618">
         - File</span></span><br><span data-ttu-id="207d0-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="207d0-620">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-620">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-621">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-621">
         - MatrixBindings</span></span><br><span data-ttu-id="207d0-622">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-622">
         - MatrixCoercion</span></span><br><span data-ttu-id="207d0-623">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-623">
         - OoxmlCoercion</span></span><br><span data-ttu-id="207d0-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-624">
         - PdfFile</span></span><br><span data-ttu-id="207d0-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-625">
         - Selection</span></span><br><span data-ttu-id="207d0-626">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-626">
         - Settings</span></span><br><span data-ttu-id="207d0-627">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-627">
         - TableBindings</span></span><br><span data-ttu-id="207d0-628">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-628">
         - TableCoercion</span></span><br><span data-ttu-id="207d0-629">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-629">
         - TextBindings</span></span><br><span data-ttu-id="207d0-630">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-630">
         - TextCoercion</span></span><br><span data-ttu-id="207d0-631">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="207d0-631">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-632">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="207d0-632">Office for Mac</span></span><br><span data-ttu-id="207d0-633">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-633">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="207d0-634">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-634">- TaskPane</span></span><br><span data-ttu-id="207d0-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-636">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-636">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="207d0-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="207d0-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="207d0-639">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="207d0-639">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="207d0-640">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-640">- BindingEvents</span></span><br><span data-ttu-id="207d0-641">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-641">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-642">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="207d0-642">
         - CustomXmlParts</span></span><br><span data-ttu-id="207d0-643">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-643">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-644">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-644">
         - File</span></span><br><span data-ttu-id="207d0-645">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-645">
         - HtmlCoercion</span></span><br><span data-ttu-id="207d0-646">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-646">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-647">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-647">
         - MatrixBindings</span></span><br><span data-ttu-id="207d0-648">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-648">
         - MatrixCoercion</span></span><br><span data-ttu-id="207d0-649">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-649">
         - OoxmlCoercion</span></span><br><span data-ttu-id="207d0-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-650">
         - PdfFile</span></span><br><span data-ttu-id="207d0-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-651">
         - Selection</span></span><br><span data-ttu-id="207d0-652">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-652">
         - Settings</span></span><br><span data-ttu-id="207d0-653">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-653">
         - TableBindings</span></span><br><span data-ttu-id="207d0-654">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-654">
         - TableCoercion</span></span><br><span data-ttu-id="207d0-655">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-655">
         - TextBindings</span></span><br><span data-ttu-id="207d0-656">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-656">
         - TextCoercion</span></span><br><span data-ttu-id="207d0-657">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="207d0-657">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-658">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="207d0-658">Office 2019 for Mac</span></span><br><span data-ttu-id="207d0-659">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-659">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-660">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-660">- TaskPane</span></span><br><span data-ttu-id="207d0-661">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-661">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-662">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-662">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="207d0-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="207d0-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="207d0-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="207d0-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="207d0-665">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="207d0-665">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="207d0-666">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-666">- BindingEvents</span></span><br><span data-ttu-id="207d0-667">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-667">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-668">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="207d0-668">
         - CustomXmlParts</span></span><br><span data-ttu-id="207d0-669">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-669">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-670">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-670">
         - File</span></span><br><span data-ttu-id="207d0-671">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-671">
         - HtmlCoercion</span></span><br><span data-ttu-id="207d0-672">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-672">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-673">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-673">
         - MatrixBindings</span></span><br><span data-ttu-id="207d0-674">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-674">
         - MatrixCoercion</span></span><br><span data-ttu-id="207d0-675">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-675">
         - OoxmlCoercion</span></span><br><span data-ttu-id="207d0-676">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-676">
         - PdfFile</span></span><br><span data-ttu-id="207d0-677">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-677">
         - Selection</span></span><br><span data-ttu-id="207d0-678">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-678">
         - Settings</span></span><br><span data-ttu-id="207d0-679">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-679">
         - TableBindings</span></span><br><span data-ttu-id="207d0-680">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-680">
         - TableCoercion</span></span><br><span data-ttu-id="207d0-681">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-681">
         - TextBindings</span></span><br><span data-ttu-id="207d0-682">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-682">
         - TextCoercion</span></span><br><span data-ttu-id="207d0-683">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="207d0-683">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-684">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="207d0-684">Office 2016 for Mac</span></span><br><span data-ttu-id="207d0-685">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-685">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-686">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="207d0-687">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-687">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="207d0-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="207d0-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="207d0-689">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-689">- BindingEvents</span></span><br><span data-ttu-id="207d0-690">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-690">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-691">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="207d0-691">
         - CustomXmlParts</span></span><br><span data-ttu-id="207d0-692">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-692">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-693">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-693">
         - File</span></span><br><span data-ttu-id="207d0-694">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-694">
         - HtmlCoercion</span></span><br><span data-ttu-id="207d0-695">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-695">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-696">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-696">
         - MatrixBindings</span></span><br><span data-ttu-id="207d0-697">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-697">
         - MatrixCoercion</span></span><br><span data-ttu-id="207d0-698">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-698">
         - OoxmlCoercion</span></span><br><span data-ttu-id="207d0-699">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-699">
         - PdfFile</span></span><br><span data-ttu-id="207d0-700">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-700">
         - Selection</span></span><br><span data-ttu-id="207d0-701">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-701">
         - Settings</span></span><br><span data-ttu-id="207d0-702">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-702">
         - TableBindings</span></span><br><span data-ttu-id="207d0-703">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-703">
         - TableCoercion</span></span><br><span data-ttu-id="207d0-704">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="207d0-704">
         - TextBindings</span></span><br><span data-ttu-id="207d0-705">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-705">
         - TextCoercion</span></span><br><span data-ttu-id="207d0-706">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="207d0-706">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="207d0-707">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="207d0-707">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="207d0-708">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="207d0-708">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="207d0-709">平台</span><span class="sxs-lookup"><span data-stu-id="207d0-709">Platform</span></span></th>
    <th><span data-ttu-id="207d0-710">扩展点</span><span class="sxs-lookup"><span data-stu-id="207d0-710">Extension points</span></span></th>
    <th><span data-ttu-id="207d0-711">API 要求集</span><span class="sxs-lookup"><span data-stu-id="207d0-711">API requirement sets</span></span></th>
    <th><span data-ttu-id="207d0-712"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="207d0-712"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-713">Office Online</span><span class="sxs-lookup"><span data-stu-id="207d0-713">Office Online</span></span></td>
    <td> <span data-ttu-id="207d0-714">- 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-714">- Content</span></span><br><span data-ttu-id="207d0-715">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-715">
         - TaskPane</span></span><br><span data-ttu-id="207d0-716">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-716">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-717">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-717">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="207d0-718">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="207d0-718">- ActiveView</span></span><br><span data-ttu-id="207d0-719">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-719">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-720">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-720">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-721">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-721">
         - File</span></span><br><span data-ttu-id="207d0-722">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-722">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-723">
         - PdfFile</span></span><br><span data-ttu-id="207d0-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-724">
         - Selection</span></span><br><span data-ttu-id="207d0-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-725">
         - Settings</span></span><br><span data-ttu-id="207d0-726">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-726">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-727">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="207d0-727">Office on Windows</span></span><br><span data-ttu-id="207d0-728">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-728">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="207d0-729">- 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-729">- Content</span></span><br><span data-ttu-id="207d0-730">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-730">
         - TaskPane</span></span><br><span data-ttu-id="207d0-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="207d0-733">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="207d0-733">- ActiveView</span></span><br><span data-ttu-id="207d0-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-734">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-735">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-736">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-736">
         - File</span></span><br><span data-ttu-id="207d0-737">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-737">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-738">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-738">
         - PdfFile</span></span><br><span data-ttu-id="207d0-739">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-739">
         - Selection</span></span><br><span data-ttu-id="207d0-740">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-740">
         - Settings</span></span><br><span data-ttu-id="207d0-741">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-741">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-742">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="207d0-742">Office 2019 on Windows</span></span><br><span data-ttu-id="207d0-743">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-743">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-744">- 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-744">- Content</span></span><br><span data-ttu-id="207d0-745">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-745">
         - TaskPane</span></span><br><span data-ttu-id="207d0-746">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-746">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-747">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-747">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="207d0-748">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="207d0-748">- ActiveView</span></span><br><span data-ttu-id="207d0-749">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-749">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-750">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-750">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-751">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-751">
         - File</span></span><br><span data-ttu-id="207d0-752">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-752">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-753">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-753">
         - PdfFile</span></span><br><span data-ttu-id="207d0-754">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-754">
         - Selection</span></span><br><span data-ttu-id="207d0-755">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-755">
         - Settings</span></span><br><span data-ttu-id="207d0-756">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-756">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-757">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="207d0-757">Office 2016 on Windows</span></span><br><span data-ttu-id="207d0-758">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-758">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-759">- 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-759">- Content</span></span><br><span data-ttu-id="207d0-760">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-760">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="207d0-761">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="207d0-761">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="207d0-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="207d0-762">- ActiveView</span></span><br><span data-ttu-id="207d0-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-763">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-764">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-765">
         - File</span></span><br><span data-ttu-id="207d0-766">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-766">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-767">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-767">
         - PdfFile</span></span><br><span data-ttu-id="207d0-768">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-768">
         - Selection</span></span><br><span data-ttu-id="207d0-769">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-769">
         - Settings</span></span><br><span data-ttu-id="207d0-770">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-770">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-771">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="207d0-771">Office 2013 on Windows</span></span><br><span data-ttu-id="207d0-772">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-772">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-773">- 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-773">- Content</span></span><br><span data-ttu-id="207d0-774">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-774">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="207d0-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="207d0-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="207d0-776">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="207d0-776">- ActiveView</span></span><br><span data-ttu-id="207d0-777">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-777">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-778">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-778">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-779">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-779">
         - File</span></span><br><span data-ttu-id="207d0-780">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-780">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-781">
         - PdfFile</span></span><br><span data-ttu-id="207d0-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-782">
         - Selection</span></span><br><span data-ttu-id="207d0-783">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-783">
         - Settings</span></span><br><span data-ttu-id="207d0-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-785">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="207d0-785">Office for iPad</span></span><br><span data-ttu-id="207d0-786">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-786">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="207d0-787">- 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-787">- Content</span></span><br><span data-ttu-id="207d0-788">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="207d0-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="207d0-790">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="207d0-790">- ActiveView</span></span><br><span data-ttu-id="207d0-791">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-791">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-792">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-792">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-793">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-793">
         - File</span></span><br><span data-ttu-id="207d0-794">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-794">
         - PdfFile</span></span><br><span data-ttu-id="207d0-795">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-795">
         - Selection</span></span><br><span data-ttu-id="207d0-796">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-796">
         - Settings</span></span><br><span data-ttu-id="207d0-797">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-797">
         - TextCoercion</span></span><br><span data-ttu-id="207d0-798">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-798">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-799">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="207d0-799">Office for Mac</span></span><br><span data-ttu-id="207d0-800">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="207d0-800">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="207d0-801">- 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-801">- Content</span></span><br><span data-ttu-id="207d0-802">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-802">
         - TaskPane</span></span><br><span data-ttu-id="207d0-803">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-803">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-804">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-804">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="207d0-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="207d0-805">- ActiveView</span></span><br><span data-ttu-id="207d0-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-806">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-807">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-808">
         - File</span></span><br><span data-ttu-id="207d0-809">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-809">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-810">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-810">
         - PdfFile</span></span><br><span data-ttu-id="207d0-811">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-811">
         - Selection</span></span><br><span data-ttu-id="207d0-812">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-812">
         - Settings</span></span><br><span data-ttu-id="207d0-813">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-813">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-814">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="207d0-814">Office 2019 for Mac</span></span><br><span data-ttu-id="207d0-815">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-815">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-816">- 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-816">- Content</span></span><br><span data-ttu-id="207d0-817">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-817">
         - TaskPane</span></span><br><span data-ttu-id="207d0-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-819">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-819">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="207d0-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="207d0-820">- ActiveView</span></span><br><span data-ttu-id="207d0-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-821">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-822">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-823">
         - File</span></span><br><span data-ttu-id="207d0-824">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-824">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-825">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-825">
         - PdfFile</span></span><br><span data-ttu-id="207d0-826">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-826">
         - Selection</span></span><br><span data-ttu-id="207d0-827">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-827">
         - Settings</span></span><br><span data-ttu-id="207d0-828">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-828">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-829">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="207d0-829">Office 2016 for Mac</span></span><br><span data-ttu-id="207d0-830">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-830">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-831">- 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-831">- Content</span></span><br><span data-ttu-id="207d0-832">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-832">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="207d0-833">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="207d0-833">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="207d0-834">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="207d0-834">- ActiveView</span></span><br><span data-ttu-id="207d0-835">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="207d0-835">
         - CompressedFile</span></span><br><span data-ttu-id="207d0-836">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-836">
         - DocumentEvents</span></span><br><span data-ttu-id="207d0-837">
         - File</span><span class="sxs-lookup"><span data-stu-id="207d0-837">
         - File</span></span><br><span data-ttu-id="207d0-838">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-838">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-839">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="207d0-839">
         - PdfFile</span></span><br><span data-ttu-id="207d0-840">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-840">
         - Selection</span></span><br><span data-ttu-id="207d0-841">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-841">
         - Settings</span></span><br><span data-ttu-id="207d0-842">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-842">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="207d0-843">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="207d0-843">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="207d0-844">OneNote</span><span class="sxs-lookup"><span data-stu-id="207d0-844">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="207d0-845">平台</span><span class="sxs-lookup"><span data-stu-id="207d0-845">Platform</span></span></th>
    <th><span data-ttu-id="207d0-846">扩展点</span><span class="sxs-lookup"><span data-stu-id="207d0-846">Extension points</span></span></th>
    <th><span data-ttu-id="207d0-847">API 要求集</span><span class="sxs-lookup"><span data-stu-id="207d0-847">API requirement sets</span></span></th>
    <th><span data-ttu-id="207d0-848"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="207d0-848"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-849">Office Online</span><span class="sxs-lookup"><span data-stu-id="207d0-849">Office Online</span></span></td>
    <td> <span data-ttu-id="207d0-850">- 内容</span><span class="sxs-lookup"><span data-stu-id="207d0-850">- Content</span></span><br><span data-ttu-id="207d0-851">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-851">
         - TaskPane</span></span><br><span data-ttu-id="207d0-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="207d0-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="207d0-853">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-853">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="207d0-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="207d0-855">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="207d0-855">- DocumentEvents</span></span><br><span data-ttu-id="207d0-856">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-856">
         - HtmlCoercion</span></span><br><span data-ttu-id="207d0-857">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-857">
         - ImageCoercion</span></span><br><span data-ttu-id="207d0-858">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="207d0-858">
         - Settings</span></span><br><span data-ttu-id="207d0-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-859">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="207d0-860">项目</span><span class="sxs-lookup"><span data-stu-id="207d0-860">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="207d0-861">平台</span><span class="sxs-lookup"><span data-stu-id="207d0-861">Platform</span></span></th>
    <th><span data-ttu-id="207d0-862">扩展点</span><span class="sxs-lookup"><span data-stu-id="207d0-862">Extension points</span></span></th>
    <th><span data-ttu-id="207d0-863">API 要求集</span><span class="sxs-lookup"><span data-stu-id="207d0-863">API requirement sets</span></span></th>
    <th><span data-ttu-id="207d0-864"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="207d0-864"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-865">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="207d0-865">Office 2019 on Windows</span></span><br><span data-ttu-id="207d0-866">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-866">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-867">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-867">- TaskPane</span></span></td>
    <td> <span data-ttu-id="207d0-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="207d0-869">- Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-869">- Selection</span></span><br><span data-ttu-id="207d0-870">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-870">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-871">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="207d0-871">Office 2016 on Windows</span></span><br><span data-ttu-id="207d0-872">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-872">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-873">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-873">- TaskPane</span></span></td>
    <td> <span data-ttu-id="207d0-874">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-874">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="207d0-875">- Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-875">- Selection</span></span><br><span data-ttu-id="207d0-876">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-876">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="207d0-877">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="207d0-877">Office 2013 on Windows</span></span><br><span data-ttu-id="207d0-878">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="207d0-878">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="207d0-879">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="207d0-879">- TaskPane</span></span></td>
    <td> <span data-ttu-id="207d0-880">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="207d0-880">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="207d0-881">- Selection</span><span class="sxs-lookup"><span data-stu-id="207d0-881">- Selection</span></span><br><span data-ttu-id="207d0-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="207d0-882">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="207d0-883">另请参阅</span><span class="sxs-lookup"><span data-stu-id="207d0-883">See also</span></span>

- [<span data-ttu-id="207d0-884">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="207d0-884">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="207d0-885">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="207d0-885">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="207d0-886">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="207d0-886">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="207d0-887">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="207d0-887">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="207d0-888">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="207d0-888">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="207d0-889">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="207d0-889">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="207d0-890">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="207d0-890">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="207d0-891">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="207d0-891">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="207d0-892">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="207d0-892">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="207d0-893">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="207d0-893">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="207d0-894">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="207d0-894">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)

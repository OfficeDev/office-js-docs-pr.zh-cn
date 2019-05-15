---
title: Office 外接程序主机和平台可用性
description: Excel、Word、Outlook、PowerPoint、OneNote 和项目支持的要求集。
ms.date: 05/08/2019
localization_priority: Priority
ms.openlocfilehash: 19f2fa7f744345823c2700b04524ec20705035a8
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952367"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="18941-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="18941-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="18941-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="18941-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="18941-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="18941-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="18941-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="18941-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="18941-108">Excel</span><span class="sxs-lookup"><span data-stu-id="18941-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="18941-109">平台</span><span class="sxs-lookup"><span data-stu-id="18941-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="18941-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="18941-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="18941-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="18941-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="18941-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="18941-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="18941-113">Office Online</span></span></td>
    <td> <span data-ttu-id="18941-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-114">- TaskPane</span></span><br><span data-ttu-id="18941-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="18941-115">
        - Content</span></span><br><span data-ttu-id="18941-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="18941-116">
        -Custom Functions</span></span><br><span data-ttu-id="18941-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="18941-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="18941-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="18941-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="18941-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="18941-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="18941-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="18941-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="18941-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="18941-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="18941-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="18941-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="18941-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="18941-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="18941-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="18941-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="18941-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="18941-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="18941-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-128">
        - BindingEvents</span></span><br><span data-ttu-id="18941-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-129">
        - CompressedFile</span></span><br><span data-ttu-id="18941-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-130">
        - DocumentEvents</span></span><br><span data-ttu-id="18941-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="18941-131">
        - File</span></span><br><span data-ttu-id="18941-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-132">
        - MatrixBindings</span></span><br><span data-ttu-id="18941-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="18941-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-134">
        - Selection</span></span><br><span data-ttu-id="18941-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-135">
        - Settings</span></span><br><span data-ttu-id="18941-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-136">
        - TableBindings</span></span><br><span data-ttu-id="18941-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-137">
        - TableCoercion</span></span><br><span data-ttu-id="18941-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-138">
        - TextBindings</span></span><br><span data-ttu-id="18941-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-140">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="18941-140">Office apps on Windows</span></span><br><span data-ttu-id="18941-141">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-141">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="18941-142">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-142">- TaskPane</span></span><br><span data-ttu-id="18941-143">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="18941-143">
        - Content</span></span><br><span data-ttu-id="18941-144">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="18941-144">
        -Custom Functions</span></span><br><span data-ttu-id="18941-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="18941-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="18941-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="18941-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="18941-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="18941-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="18941-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="18941-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="18941-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="18941-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="18941-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="18941-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="18941-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="18941-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="18941-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="18941-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="18941-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="18941-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="18941-156">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-156">
        - BindingEvents</span></span><br><span data-ttu-id="18941-157">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-157">
        - CompressedFile</span></span><br><span data-ttu-id="18941-158">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-158">
        - DocumentEvents</span></span><br><span data-ttu-id="18941-159">
        - File</span><span class="sxs-lookup"><span data-stu-id="18941-159">
        - File</span></span><br><span data-ttu-id="18941-160">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-160">
        - MatrixBindings</span></span><br><span data-ttu-id="18941-161">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-161">
        - MatrixCoercion</span></span><br><span data-ttu-id="18941-162">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-162">
        - Selection</span></span><br><span data-ttu-id="18941-163">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-163">
        - Settings</span></span><br><span data-ttu-id="18941-164">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-164">
        - TableBindings</span></span><br><span data-ttu-id="18941-165">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-165">
        - TableCoercion</span></span><br><span data-ttu-id="18941-166">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-166">
        - TextBindings</span></span><br><span data-ttu-id="18941-167">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-167">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-168">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="18941-168">Office 2019 for Windows</span></span><br><span data-ttu-id="18941-169">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-169">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="18941-170">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-170">- TaskPane</span></span><br><span data-ttu-id="18941-171">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="18941-171">
        - Content</span></span><br><span data-ttu-id="18941-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="18941-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="18941-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="18941-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="18941-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="18941-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="18941-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="18941-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="18941-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="18941-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="18941-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="18941-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="18941-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="18941-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="18941-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="18941-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-182">- BindingEvents</span></span><br><span data-ttu-id="18941-183">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-183">
        - CompressedFile</span></span><br><span data-ttu-id="18941-184">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-184">
        - DocumentEvents</span></span><br><span data-ttu-id="18941-185">
        - File</span><span class="sxs-lookup"><span data-stu-id="18941-185">
        - File</span></span><br><span data-ttu-id="18941-186">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-186">
        - ImageCoercion</span></span><br><span data-ttu-id="18941-187">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-187">
        - MatrixBindings</span></span><br><span data-ttu-id="18941-188">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-188">
        - MatrixCoercion</span></span><br><span data-ttu-id="18941-189">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-189">
        - Selection</span></span><br><span data-ttu-id="18941-190">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-190">
        - Settings</span></span><br><span data-ttu-id="18941-191">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-191">
        - TableBindings</span></span><br><span data-ttu-id="18941-192">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-192">
        - TableCoercion</span></span><br><span data-ttu-id="18941-193">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-193">
        - TextBindings</span></span><br><span data-ttu-id="18941-194">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-194">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-195">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="18941-195">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="18941-196">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-196">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="18941-197">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-197">- TaskPane</span></span><br><span data-ttu-id="18941-198">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="18941-198">
        - Content</span></span></td>
    <td><span data-ttu-id="18941-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="18941-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="18941-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="18941-201">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-201">- BindingEvents</span></span><br><span data-ttu-id="18941-202">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-202">
        - CompressedFile</span></span><br><span data-ttu-id="18941-203">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-203">
        - DocumentEvents</span></span><br><span data-ttu-id="18941-204">
        - File</span><span class="sxs-lookup"><span data-stu-id="18941-204">
        - File</span></span><br><span data-ttu-id="18941-205">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-205">
        - ImageCoercion</span></span><br><span data-ttu-id="18941-206">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-206">
        - MatrixBindings</span></span><br><span data-ttu-id="18941-207">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-207">
        - MatrixCoercion</span></span><br><span data-ttu-id="18941-208">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-208">
        - Selection</span></span><br><span data-ttu-id="18941-209">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-209">
        - Settings</span></span><br><span data-ttu-id="18941-210">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-210">
        - TableBindings</span></span><br><span data-ttu-id="18941-211">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-211">
        - TableCoercion</span></span><br><span data-ttu-id="18941-212">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-212">
        - TextBindings</span></span><br><span data-ttu-id="18941-213">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-213">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-214">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="18941-214">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="18941-215">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-215">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="18941-216">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-216">
        - TaskPane</span></span><br><span data-ttu-id="18941-217">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="18941-217">
        - Content</span></span></td>
    <td>  <span data-ttu-id="18941-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="18941-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="18941-219">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-219">
        - BindingEvents</span></span><br><span data-ttu-id="18941-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-220">
        - CompressedFile</span></span><br><span data-ttu-id="18941-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-221">
        - DocumentEvents</span></span><br><span data-ttu-id="18941-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="18941-222">
        - File</span></span><br><span data-ttu-id="18941-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-223">
        - ImageCoercion</span></span><br><span data-ttu-id="18941-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-224">
        - MatrixBindings</span></span><br><span data-ttu-id="18941-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="18941-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-226">
        - Selection</span></span><br><span data-ttu-id="18941-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-227">
        - Settings</span></span><br><span data-ttu-id="18941-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-228">
        - TableBindings</span></span><br><span data-ttu-id="18941-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-229">
        - TableCoercion</span></span><br><span data-ttu-id="18941-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-230">
        - TextBindings</span></span><br><span data-ttu-id="18941-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-232">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="18941-232">Office for iPad</span></span><br><span data-ttu-id="18941-233">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-233">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="18941-234">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-234">- TaskPane</span></span><br><span data-ttu-id="18941-235">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="18941-235">
        - Content</span></span><br><span data-ttu-id="18941-236">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="18941-236">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="18941-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="18941-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="18941-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="18941-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="18941-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="18941-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="18941-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="18941-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="18941-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="18941-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="18941-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="18941-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="18941-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="18941-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="18941-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="18941-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="18941-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-247">- BindingEvents</span></span><br><span data-ttu-id="18941-248">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-248">
        - DocumentEvents</span></span><br><span data-ttu-id="18941-249">
        - File</span><span class="sxs-lookup"><span data-stu-id="18941-249">
        - File</span></span><br><span data-ttu-id="18941-250">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-250">
        - ImageCoercion</span></span><br><span data-ttu-id="18941-251">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-251">
        - MatrixBindings</span></span><br><span data-ttu-id="18941-252">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-252">
        - MatrixCoercion</span></span><br><span data-ttu-id="18941-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-253">
        - Selection</span></span><br><span data-ttu-id="18941-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-254">
        - Settings</span></span><br><span data-ttu-id="18941-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-255">
        - TableBindings</span></span><br><span data-ttu-id="18941-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-256">
        - TableCoercion</span></span><br><span data-ttu-id="18941-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-257">
        - TextBindings</span></span><br><span data-ttu-id="18941-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-259">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="18941-259">Office for Mac</span></span><br><span data-ttu-id="18941-260">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-260">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="18941-261">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-261">- TaskPane</span></span><br><span data-ttu-id="18941-262">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="18941-262">
        - Content</span></span><br><span data-ttu-id="18941-263">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="18941-263">
        -Custom Functions</span></span><br><span data-ttu-id="18941-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="18941-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="18941-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="18941-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="18941-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="18941-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="18941-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="18941-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="18941-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="18941-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="18941-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="18941-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="18941-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="18941-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="18941-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="18941-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="18941-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="18941-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-275">- BindingEvents</span></span><br><span data-ttu-id="18941-276">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-276">
        - CompressedFile</span></span><br><span data-ttu-id="18941-277">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-277">
        - DocumentEvents</span></span><br><span data-ttu-id="18941-278">
        - File</span><span class="sxs-lookup"><span data-stu-id="18941-278">
        - File</span></span><br><span data-ttu-id="18941-279">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-279">
        - ImageCoercion</span></span><br><span data-ttu-id="18941-280">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-280">
        - MatrixBindings</span></span><br><span data-ttu-id="18941-281">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-281">
        - MatrixCoercion</span></span><br><span data-ttu-id="18941-282">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-282">
        - PdfFile</span></span><br><span data-ttu-id="18941-283">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-283">
        - Selection</span></span><br><span data-ttu-id="18941-284">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-284">
        - Settings</span></span><br><span data-ttu-id="18941-285">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-285">
        - TableBindings</span></span><br><span data-ttu-id="18941-286">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-286">
        - TableCoercion</span></span><br><span data-ttu-id="18941-287">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-287">
        - TextBindings</span></span><br><span data-ttu-id="18941-288">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-288">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-289">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="18941-289">Office 2019 for Mac</span></span><br><span data-ttu-id="18941-290">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-290">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="18941-291">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-291">- TaskPane</span></span><br><span data-ttu-id="18941-292">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="18941-292">
        - Content</span></span><br><span data-ttu-id="18941-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="18941-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="18941-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="18941-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="18941-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="18941-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="18941-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="18941-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="18941-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="18941-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="18941-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="18941-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="18941-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="18941-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="18941-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="18941-303">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-303">- BindingEvents</span></span><br><span data-ttu-id="18941-304">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-304">
        - CompressedFile</span></span><br><span data-ttu-id="18941-305">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-305">
        - DocumentEvents</span></span><br><span data-ttu-id="18941-306">
        - File</span><span class="sxs-lookup"><span data-stu-id="18941-306">
        - File</span></span><br><span data-ttu-id="18941-307">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-307">
        - ImageCoercion</span></span><br><span data-ttu-id="18941-308">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-308">
        - MatrixBindings</span></span><br><span data-ttu-id="18941-309">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-309">
        - MatrixCoercion</span></span><br><span data-ttu-id="18941-310">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-310">
        - PdfFile</span></span><br><span data-ttu-id="18941-311">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-311">
        - Selection</span></span><br><span data-ttu-id="18941-312">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-312">
        - Settings</span></span><br><span data-ttu-id="18941-313">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-313">
        - TableBindings</span></span><br><span data-ttu-id="18941-314">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-314">
        - TableCoercion</span></span><br><span data-ttu-id="18941-315">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-315">
        - TextBindings</span></span><br><span data-ttu-id="18941-316">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-316">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-317">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="18941-317">Office 2016 for Mac</span></span><br><span data-ttu-id="18941-318">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-318">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="18941-319">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-319">- TaskPane</span></span><br><span data-ttu-id="18941-320">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="18941-320">
        - Content</span></span></td>
    <td><span data-ttu-id="18941-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="18941-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="18941-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="18941-323">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-323">- BindingEvents</span></span><br><span data-ttu-id="18941-324">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-324">
        - CompressedFile</span></span><br><span data-ttu-id="18941-325">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-325">
        - DocumentEvents</span></span><br><span data-ttu-id="18941-326">
        - File</span><span class="sxs-lookup"><span data-stu-id="18941-326">
        - File</span></span><br><span data-ttu-id="18941-327">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-327">
        - ImageCoercion</span></span><br><span data-ttu-id="18941-328">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-328">
        - MatrixBindings</span></span><br><span data-ttu-id="18941-329">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-329">
        - MatrixCoercion</span></span><br><span data-ttu-id="18941-330">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-330">
        - PdfFile</span></span><br><span data-ttu-id="18941-331">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-331">
        - Selection</span></span><br><span data-ttu-id="18941-332">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-332">
        - Settings</span></span><br><span data-ttu-id="18941-333">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-333">
        - TableBindings</span></span><br><span data-ttu-id="18941-334">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-334">
        - TableCoercion</span></span><br><span data-ttu-id="18941-335">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-335">
        - TextBindings</span></span><br><span data-ttu-id="18941-336">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-336">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="18941-337">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="18941-337">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="18941-338">自定义函数</span><span class="sxs-lookup"><span data-stu-id="18941-338">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="18941-339">平台</span><span class="sxs-lookup"><span data-stu-id="18941-339">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="18941-340">扩展点</span><span class="sxs-lookup"><span data-stu-id="18941-340">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="18941-341">API 要求集</span><span class="sxs-lookup"><span data-stu-id="18941-341">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="18941-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="18941-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-343">Office Online</span><span class="sxs-lookup"><span data-stu-id="18941-343">Office Online</span></span></td>
    <td><span data-ttu-id="18941-344">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="18941-344">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="18941-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-346">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="18941-346">Office apps on Windows</span></span><br><span data-ttu-id="18941-347">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-347">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="18941-348">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="18941-348">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="18941-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-350">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="18941-350">Office for iPad</span></span><br><span data-ttu-id="18941-351">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-351">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="18941-352">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="18941-352">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="18941-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-354">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="18941-354">Office for Mac</span></span><br><span data-ttu-id="18941-355">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-355">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="18941-356">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="18941-356">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="18941-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="18941-358">Outlook</span><span class="sxs-lookup"><span data-stu-id="18941-358">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="18941-359">平台</span><span class="sxs-lookup"><span data-stu-id="18941-359">Platform</span></span></th>
    <th><span data-ttu-id="18941-360">扩展点</span><span class="sxs-lookup"><span data-stu-id="18941-360">Extension points</span></span></th>
    <th><span data-ttu-id="18941-361">API 要求集</span><span class="sxs-lookup"><span data-stu-id="18941-361">API requirement sets</span></span></th>
    <th><span data-ttu-id="18941-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="18941-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-363">Office Online</span><span class="sxs-lookup"><span data-stu-id="18941-363">Office Online</span></span></td>
    <td> <span data-ttu-id="18941-364">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="18941-364">- Mail Read</span></span><br><span data-ttu-id="18941-365">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="18941-365">
      - Mail Compose</span></span><br><span data-ttu-id="18941-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="18941-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="18941-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="18941-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="18941-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="18941-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="18941-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="18941-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="18941-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="18941-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="18941-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="18941-374">不可用</span><span class="sxs-lookup"><span data-stu-id="18941-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-375">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="18941-375">Office apps on Windows</span></span><br><span data-ttu-id="18941-376">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-376">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="18941-377">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="18941-377">- Mail Read</span></span><br><span data-ttu-id="18941-378">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="18941-378">
      - Mail Compose</span></span><br><span data-ttu-id="18941-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="18941-380">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="18941-380">
      - Modules</span></span></td>
    <td> <span data-ttu-id="18941-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="18941-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="18941-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="18941-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="18941-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="18941-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="18941-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="18941-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="18941-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="18941-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="18941-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="18941-388">不可用</span><span class="sxs-lookup"><span data-stu-id="18941-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-389">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="18941-389">Office 2019 for Windows</span></span><br><span data-ttu-id="18941-390">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-390">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-391">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="18941-391">- Mail Read</span></span><br><span data-ttu-id="18941-392">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="18941-392">
      - Mail Compose</span></span><br><span data-ttu-id="18941-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="18941-394">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="18941-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="18941-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="18941-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="18941-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="18941-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="18941-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="18941-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="18941-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="18941-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="18941-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="18941-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="18941-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="18941-402">不可用</span><span class="sxs-lookup"><span data-stu-id="18941-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-403">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="18941-403">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="18941-404">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-405">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="18941-405">- Mail Read</span></span><br><span data-ttu-id="18941-406">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="18941-406">
      - Mail Compose</span></span><br><span data-ttu-id="18941-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="18941-408">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="18941-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="18941-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="18941-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="18941-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="18941-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="18941-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="18941-413">不可用</span><span class="sxs-lookup"><span data-stu-id="18941-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-414">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="18941-414">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="18941-415">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-416">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="18941-416">- Mail Read</span></span><br><span data-ttu-id="18941-417">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="18941-417">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="18941-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="18941-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="18941-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="18941-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="18941-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="18941-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="18941-422">不可用</span><span class="sxs-lookup"><span data-stu-id="18941-422">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-423">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="18941-423">Office for iOS</span></span><br><span data-ttu-id="18941-424">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-424">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="18941-425">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="18941-425">- Mail Read</span></span><br><span data-ttu-id="18941-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="18941-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="18941-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="18941-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="18941-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="18941-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="18941-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="18941-432">不可用</span><span class="sxs-lookup"><span data-stu-id="18941-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-433">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="18941-433">Office for Mac</span></span><br><span data-ttu-id="18941-434">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-434">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="18941-435">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="18941-435">- Mail Read</span></span><br><span data-ttu-id="18941-436">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="18941-436">
      - Mail Compose</span></span><br><span data-ttu-id="18941-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="18941-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="18941-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="18941-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="18941-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="18941-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="18941-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="18941-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="18941-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="18941-444">不可用</span><span class="sxs-lookup"><span data-stu-id="18941-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-445">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="18941-445">Office 2019 for Mac</span></span><br><span data-ttu-id="18941-446">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-446">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-447">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="18941-447">- Mail Read</span></span><br><span data-ttu-id="18941-448">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="18941-448">
      - Mail Compose</span></span><br><span data-ttu-id="18941-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="18941-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="18941-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="18941-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="18941-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="18941-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="18941-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="18941-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="18941-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="18941-456">不可用</span><span class="sxs-lookup"><span data-stu-id="18941-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-457">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="18941-457">Office 2016 for Mac</span></span><br><span data-ttu-id="18941-458">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-458">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-459">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="18941-459">- Mail Read</span></span><br><span data-ttu-id="18941-460">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="18941-460">
      - Mail Compose</span></span><br><span data-ttu-id="18941-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="18941-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="18941-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="18941-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="18941-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="18941-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="18941-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="18941-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="18941-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="18941-468">不可用</span><span class="sxs-lookup"><span data-stu-id="18941-468">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-469">Office for Android</span><span class="sxs-lookup"><span data-stu-id="18941-469">Office for Android</span></span><br><span data-ttu-id="18941-470">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-470">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="18941-471">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="18941-471">- Mail Read</span></span><br><span data-ttu-id="18941-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="18941-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="18941-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="18941-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="18941-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="18941-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="18941-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="18941-478">不可用</span><span class="sxs-lookup"><span data-stu-id="18941-478">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="18941-479">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="18941-479">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="18941-480">Word</span><span class="sxs-lookup"><span data-stu-id="18941-480">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="18941-481">平台</span><span class="sxs-lookup"><span data-stu-id="18941-481">Platform</span></span></th>
    <th><span data-ttu-id="18941-482">扩展点</span><span class="sxs-lookup"><span data-stu-id="18941-482">Extension points</span></span></th>
    <th><span data-ttu-id="18941-483">API 要求集</span><span class="sxs-lookup"><span data-stu-id="18941-483">API requirement sets</span></span></th>
    <th><span data-ttu-id="18941-484"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="18941-484"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-485">Office Online</span><span class="sxs-lookup"><span data-stu-id="18941-485">Office Online</span></span></td>
    <td> <span data-ttu-id="18941-486">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-486">- TaskPane</span></span><br><span data-ttu-id="18941-487">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-487">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-488">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-488">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="18941-489">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-489">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="18941-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="18941-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="18941-492">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-492">- BindingEvents</span></span><br><span data-ttu-id="18941-493">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="18941-493">
         - CustomXmlParts</span></span><br><span data-ttu-id="18941-494">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-494">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-495">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-495">
         - File</span></span><br><span data-ttu-id="18941-496">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-496">
         - HtmlCoercion</span></span><br><span data-ttu-id="18941-497">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-497">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-498">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-498">
         - MatrixBindings</span></span><br><span data-ttu-id="18941-499">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-499">
         - MatrixCoercion</span></span><br><span data-ttu-id="18941-500">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-500">
         - OoxmlCoercion</span></span><br><span data-ttu-id="18941-501">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-501">
         - PdfFile</span></span><br><span data-ttu-id="18941-502">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-502">
         - Selection</span></span><br><span data-ttu-id="18941-503">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-503">
         - Settings</span></span><br><span data-ttu-id="18941-504">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-504">
         - TableBindings</span></span><br><span data-ttu-id="18941-505">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-505">
         - TableCoercion</span></span><br><span data-ttu-id="18941-506">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-506">
         - TextBindings</span></span><br><span data-ttu-id="18941-507">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-507">
         - TextCoercion</span></span><br><span data-ttu-id="18941-508">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="18941-508">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-509">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="18941-509">Office apps on Windows</span></span><br><span data-ttu-id="18941-510">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-510">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="18941-511">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-511">- TaskPane</span></span><br><span data-ttu-id="18941-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-513">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-513">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="18941-514">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-514">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="18941-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="18941-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="18941-517">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-517">- BindingEvents</span></span><br><span data-ttu-id="18941-518">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-518">
         - CompressedFile</span></span><br><span data-ttu-id="18941-519">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="18941-519">
         - CustomXmlParts</span></span><br><span data-ttu-id="18941-520">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-520">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-521">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-521">
         - File</span></span><br><span data-ttu-id="18941-522">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-522">
         - HtmlCoercion</span></span><br><span data-ttu-id="18941-523">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-523">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-524">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-524">
         - MatrixBindings</span></span><br><span data-ttu-id="18941-525">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-525">
         - MatrixCoercion</span></span><br><span data-ttu-id="18941-526">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-526">
         - OoxmlCoercion</span></span><br><span data-ttu-id="18941-527">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-527">
         - PdfFile</span></span><br><span data-ttu-id="18941-528">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-528">
         - Selection</span></span><br><span data-ttu-id="18941-529">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-529">
         - Settings</span></span><br><span data-ttu-id="18941-530">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-530">
         - TableBindings</span></span><br><span data-ttu-id="18941-531">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-531">
         - TableCoercion</span></span><br><span data-ttu-id="18941-532">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-532">
         - TextBindings</span></span><br><span data-ttu-id="18941-533">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-533">
         - TextCoercion</span></span><br><span data-ttu-id="18941-534">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="18941-534">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-535">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="18941-535">Office 2019 for Windows</span></span><br><span data-ttu-id="18941-536">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-536">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-537">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-537">- TaskPane</span></span><br><span data-ttu-id="18941-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-539">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-539">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="18941-540">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-540">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="18941-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="18941-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="18941-543">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-543">- BindingEvents</span></span><br><span data-ttu-id="18941-544">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-544">
         - CompressedFile</span></span><br><span data-ttu-id="18941-545">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="18941-545">
         - CustomXmlParts</span></span><br><span data-ttu-id="18941-546">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-546">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-547">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-547">
         - File</span></span><br><span data-ttu-id="18941-548">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-548">
         - HtmlCoercion</span></span><br><span data-ttu-id="18941-549">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-549">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-550">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-550">
         - MatrixBindings</span></span><br><span data-ttu-id="18941-551">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-551">
         - MatrixCoercion</span></span><br><span data-ttu-id="18941-552">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-552">
         - OoxmlCoercion</span></span><br><span data-ttu-id="18941-553">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-553">
         - PdfFile</span></span><br><span data-ttu-id="18941-554">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-554">
         - Selection</span></span><br><span data-ttu-id="18941-555">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-555">
         - Settings</span></span><br><span data-ttu-id="18941-556">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-556">
         - TableBindings</span></span><br><span data-ttu-id="18941-557">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-557">
         - TableCoercion</span></span><br><span data-ttu-id="18941-558">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-558">
         - TextBindings</span></span><br><span data-ttu-id="18941-559">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-559">
         - TextCoercion</span></span><br><span data-ttu-id="18941-560">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="18941-560">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-561">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="18941-561">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="18941-562">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-562">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-563">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-563">- TaskPane</span></span></td>
    <td> <span data-ttu-id="18941-564">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-564">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="18941-565">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="18941-565">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="18941-566">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-566">- BindingEvents</span></span><br><span data-ttu-id="18941-567">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-567">
         - CompressedFile</span></span><br><span data-ttu-id="18941-568">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="18941-568">
         - CustomXmlParts</span></span><br><span data-ttu-id="18941-569">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-569">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-570">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-570">
         - File</span></span><br><span data-ttu-id="18941-571">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-571">
         - HtmlCoercion</span></span><br><span data-ttu-id="18941-572">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-572">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-573">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-573">
         - MatrixBindings</span></span><br><span data-ttu-id="18941-574">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-574">
         - MatrixCoercion</span></span><br><span data-ttu-id="18941-575">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-575">
         - OoxmlCoercion</span></span><br><span data-ttu-id="18941-576">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-576">
         - PdfFile</span></span><br><span data-ttu-id="18941-577">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-577">
         - Selection</span></span><br><span data-ttu-id="18941-578">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-578">
         - Settings</span></span><br><span data-ttu-id="18941-579">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-579">
         - TableBindings</span></span><br><span data-ttu-id="18941-580">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-580">
         - TableCoercion</span></span><br><span data-ttu-id="18941-581">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-581">
         - TextBindings</span></span><br><span data-ttu-id="18941-582">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-582">
         - TextCoercion</span></span><br><span data-ttu-id="18941-583">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="18941-583">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-584">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="18941-584">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="18941-585">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-585">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-586">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-586">- TaskPane</span></span></td>
    <td> <span data-ttu-id="18941-587">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="18941-587">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="18941-588">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-588">- BindingEvents</span></span><br><span data-ttu-id="18941-589">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-589">
         - CompressedFile</span></span><br><span data-ttu-id="18941-590">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="18941-590">
         - CustomXmlParts</span></span><br><span data-ttu-id="18941-591">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-591">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-592">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-592">
         - File</span></span><br><span data-ttu-id="18941-593">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-593">
         - HtmlCoercion</span></span><br><span data-ttu-id="18941-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-594">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-595">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-595">
         - MatrixBindings</span></span><br><span data-ttu-id="18941-596">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-596">
         - MatrixCoercion</span></span><br><span data-ttu-id="18941-597">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-597">
         - OoxmlCoercion</span></span><br><span data-ttu-id="18941-598">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-598">
         - PdfFile</span></span><br><span data-ttu-id="18941-599">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-599">
         - Selection</span></span><br><span data-ttu-id="18941-600">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-600">
         - Settings</span></span><br><span data-ttu-id="18941-601">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-601">
         - TableBindings</span></span><br><span data-ttu-id="18941-602">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-602">
         - TableCoercion</span></span><br><span data-ttu-id="18941-603">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-603">
         - TextBindings</span></span><br><span data-ttu-id="18941-604">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-604">
         - TextCoercion</span></span><br><span data-ttu-id="18941-605">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="18941-605">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-606">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="18941-606">Office for iPad</span></span><br><span data-ttu-id="18941-607">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-607">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="18941-608">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-608">- TaskPane</span></span></td>
    <td> <span data-ttu-id="18941-609">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-609">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="18941-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="18941-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="18941-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="18941-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="18941-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-613">- BindingEvents</span></span><br><span data-ttu-id="18941-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-614">
         - CompressedFile</span></span><br><span data-ttu-id="18941-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="18941-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="18941-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-616">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-617">
         - File</span></span><br><span data-ttu-id="18941-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="18941-619">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-619">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-620">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-620">
         - MatrixBindings</span></span><br><span data-ttu-id="18941-621">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-621">
         - MatrixCoercion</span></span><br><span data-ttu-id="18941-622">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-622">
         - OoxmlCoercion</span></span><br><span data-ttu-id="18941-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-623">
         - PdfFile</span></span><br><span data-ttu-id="18941-624">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-624">
         - Selection</span></span><br><span data-ttu-id="18941-625">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-625">
         - Settings</span></span><br><span data-ttu-id="18941-626">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-626">
         - TableBindings</span></span><br><span data-ttu-id="18941-627">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-627">
         - TableCoercion</span></span><br><span data-ttu-id="18941-628">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-628">
         - TextBindings</span></span><br><span data-ttu-id="18941-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-629">
         - TextCoercion</span></span><br><span data-ttu-id="18941-630">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="18941-630">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-631">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="18941-631">Office for Mac</span></span><br><span data-ttu-id="18941-632">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-632">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="18941-633">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-633">- TaskPane</span></span><br><span data-ttu-id="18941-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="18941-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="18941-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="18941-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="18941-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="18941-639">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-639">- BindingEvents</span></span><br><span data-ttu-id="18941-640">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-640">
         - CompressedFile</span></span><br><span data-ttu-id="18941-641">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="18941-641">
         - CustomXmlParts</span></span><br><span data-ttu-id="18941-642">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-642">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-643">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-643">
         - File</span></span><br><span data-ttu-id="18941-644">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-644">
         - HtmlCoercion</span></span><br><span data-ttu-id="18941-645">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-645">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-646">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-646">
         - MatrixBindings</span></span><br><span data-ttu-id="18941-647">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-647">
         - MatrixCoercion</span></span><br><span data-ttu-id="18941-648">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-648">
         - OoxmlCoercion</span></span><br><span data-ttu-id="18941-649">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-649">
         - PdfFile</span></span><br><span data-ttu-id="18941-650">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-650">
         - Selection</span></span><br><span data-ttu-id="18941-651">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-651">
         - Settings</span></span><br><span data-ttu-id="18941-652">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-652">
         - TableBindings</span></span><br><span data-ttu-id="18941-653">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-653">
         - TableCoercion</span></span><br><span data-ttu-id="18941-654">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-654">
         - TextBindings</span></span><br><span data-ttu-id="18941-655">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-655">
         - TextCoercion</span></span><br><span data-ttu-id="18941-656">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="18941-656">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-657">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="18941-657">Office 2019 for Mac</span></span><br><span data-ttu-id="18941-658">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-658">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-659">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-659">- TaskPane</span></span><br><span data-ttu-id="18941-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-661">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-661">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="18941-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="18941-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="18941-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="18941-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="18941-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="18941-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="18941-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-665">- BindingEvents</span></span><br><span data-ttu-id="18941-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-666">
         - CompressedFile</span></span><br><span data-ttu-id="18941-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="18941-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="18941-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-668">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-669">
         - File</span></span><br><span data-ttu-id="18941-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="18941-671">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-671">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-672">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-672">
         - MatrixBindings</span></span><br><span data-ttu-id="18941-673">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-673">
         - MatrixCoercion</span></span><br><span data-ttu-id="18941-674">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-674">
         - OoxmlCoercion</span></span><br><span data-ttu-id="18941-675">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-675">
         - PdfFile</span></span><br><span data-ttu-id="18941-676">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-676">
         - Selection</span></span><br><span data-ttu-id="18941-677">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-677">
         - Settings</span></span><br><span data-ttu-id="18941-678">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-678">
         - TableBindings</span></span><br><span data-ttu-id="18941-679">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-679">
         - TableCoercion</span></span><br><span data-ttu-id="18941-680">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-680">
         - TextBindings</span></span><br><span data-ttu-id="18941-681">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-681">
         - TextCoercion</span></span><br><span data-ttu-id="18941-682">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="18941-682">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-683">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="18941-683">Office 2016 for Mac</span></span><br><span data-ttu-id="18941-684">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-684">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-685">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-685">- TaskPane</span></span></td>
    <td> <span data-ttu-id="18941-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="18941-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="18941-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="18941-688">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="18941-688">- BindingEvents</span></span><br><span data-ttu-id="18941-689">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-689">
         - CompressedFile</span></span><br><span data-ttu-id="18941-690">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="18941-690">
         - CustomXmlParts</span></span><br><span data-ttu-id="18941-691">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-691">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-692">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-692">
         - File</span></span><br><span data-ttu-id="18941-693">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-693">
         - HtmlCoercion</span></span><br><span data-ttu-id="18941-694">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-694">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-695">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="18941-695">
         - MatrixBindings</span></span><br><span data-ttu-id="18941-696">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-696">
         - MatrixCoercion</span></span><br><span data-ttu-id="18941-697">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-697">
         - OoxmlCoercion</span></span><br><span data-ttu-id="18941-698">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-698">
         - PdfFile</span></span><br><span data-ttu-id="18941-699">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-699">
         - Selection</span></span><br><span data-ttu-id="18941-700">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-700">
         - Settings</span></span><br><span data-ttu-id="18941-701">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="18941-701">
         - TableBindings</span></span><br><span data-ttu-id="18941-702">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-702">
         - TableCoercion</span></span><br><span data-ttu-id="18941-703">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="18941-703">
         - TextBindings</span></span><br><span data-ttu-id="18941-704">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-704">
         - TextCoercion</span></span><br><span data-ttu-id="18941-705">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="18941-705">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="18941-706">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="18941-706">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="18941-707">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="18941-707">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="18941-708">平台</span><span class="sxs-lookup"><span data-stu-id="18941-708">Platform</span></span></th>
    <th><span data-ttu-id="18941-709">扩展点</span><span class="sxs-lookup"><span data-stu-id="18941-709">Extension points</span></span></th>
    <th><span data-ttu-id="18941-710">API 要求集</span><span class="sxs-lookup"><span data-stu-id="18941-710">API requirement sets</span></span></th>
    <th><span data-ttu-id="18941-711"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="18941-711"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-712">Office Online</span><span class="sxs-lookup"><span data-stu-id="18941-712">Office Online</span></span></td>
    <td> <span data-ttu-id="18941-713">- 内容</span><span class="sxs-lookup"><span data-stu-id="18941-713">- Content</span></span><br><span data-ttu-id="18941-714">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-714">
         - TaskPane</span></span><br><span data-ttu-id="18941-715">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-715">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-716">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-716">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="18941-717">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="18941-717">- ActiveView</span></span><br><span data-ttu-id="18941-718">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-718">
         - CompressedFile</span></span><br><span data-ttu-id="18941-719">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-719">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-720">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-720">
         - File</span></span><br><span data-ttu-id="18941-721">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-721">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-722">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-722">
         - PdfFile</span></span><br><span data-ttu-id="18941-723">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-723">
         - Selection</span></span><br><span data-ttu-id="18941-724">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-724">
         - Settings</span></span><br><span data-ttu-id="18941-725">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-725">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-726">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="18941-726">Office apps on Windows</span></span><br><span data-ttu-id="18941-727">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-727">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="18941-728">- 内容</span><span class="sxs-lookup"><span data-stu-id="18941-728">- Content</span></span><br><span data-ttu-id="18941-729">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-729">
         - TaskPane</span></span><br><span data-ttu-id="18941-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="18941-732">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="18941-732">- ActiveView</span></span><br><span data-ttu-id="18941-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-733">
         - CompressedFile</span></span><br><span data-ttu-id="18941-734">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-734">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-735">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-735">
         - File</span></span><br><span data-ttu-id="18941-736">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-736">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-737">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-737">
         - PdfFile</span></span><br><span data-ttu-id="18941-738">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-738">
         - Selection</span></span><br><span data-ttu-id="18941-739">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-739">
         - Settings</span></span><br><span data-ttu-id="18941-740">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-740">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-741">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="18941-741">Office 2019 for Windows</span></span><br><span data-ttu-id="18941-742">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-742">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-743">- 内容</span><span class="sxs-lookup"><span data-stu-id="18941-743">- Content</span></span><br><span data-ttu-id="18941-744">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-744">
         - TaskPane</span></span><br><span data-ttu-id="18941-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="18941-747">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="18941-747">- ActiveView</span></span><br><span data-ttu-id="18941-748">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-748">
         - CompressedFile</span></span><br><span data-ttu-id="18941-749">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-749">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-750">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-750">
         - File</span></span><br><span data-ttu-id="18941-751">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-751">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-752">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-752">
         - PdfFile</span></span><br><span data-ttu-id="18941-753">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-753">
         - Selection</span></span><br><span data-ttu-id="18941-754">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-754">
         - Settings</span></span><br><span data-ttu-id="18941-755">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-755">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-756">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="18941-756">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="18941-757">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-757">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-758">- 内容</span><span class="sxs-lookup"><span data-stu-id="18941-758">- Content</span></span><br><span data-ttu-id="18941-759">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-759">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="18941-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="18941-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="18941-761">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="18941-761">- ActiveView</span></span><br><span data-ttu-id="18941-762">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-762">
         - CompressedFile</span></span><br><span data-ttu-id="18941-763">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-763">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-764">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-764">
         - File</span></span><br><span data-ttu-id="18941-765">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-765">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-766">
         - PdfFile</span></span><br><span data-ttu-id="18941-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-767">
         - Selection</span></span><br><span data-ttu-id="18941-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-768">
         - Settings</span></span><br><span data-ttu-id="18941-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-770">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="18941-770">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="18941-771">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-772">- 内容</span><span class="sxs-lookup"><span data-stu-id="18941-772">- Content</span></span><br><span data-ttu-id="18941-773">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-773">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="18941-774">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="18941-774">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="18941-775">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="18941-775">- ActiveView</span></span><br><span data-ttu-id="18941-776">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-776">
         - CompressedFile</span></span><br><span data-ttu-id="18941-777">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-777">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-778">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-778">
         - File</span></span><br><span data-ttu-id="18941-779">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-779">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-780">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-780">
         - PdfFile</span></span><br><span data-ttu-id="18941-781">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-781">
         - Selection</span></span><br><span data-ttu-id="18941-782">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-782">
         - Settings</span></span><br><span data-ttu-id="18941-783">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-783">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-784">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="18941-784">Office for iPad</span></span><br><span data-ttu-id="18941-785">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-785">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="18941-786">- 内容</span><span class="sxs-lookup"><span data-stu-id="18941-786">- Content</span></span><br><span data-ttu-id="18941-787">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-787">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="18941-788">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-788">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="18941-789">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="18941-789">- ActiveView</span></span><br><span data-ttu-id="18941-790">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-790">
         - CompressedFile</span></span><br><span data-ttu-id="18941-791">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-791">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-792">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-792">
         - File</span></span><br><span data-ttu-id="18941-793">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-793">
         - PdfFile</span></span><br><span data-ttu-id="18941-794">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-794">
         - Selection</span></span><br><span data-ttu-id="18941-795">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-795">
         - Settings</span></span><br><span data-ttu-id="18941-796">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-796">
         - TextCoercion</span></span><br><span data-ttu-id="18941-797">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-797">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-798">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="18941-798">Office for Mac</span></span><br><span data-ttu-id="18941-799">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="18941-799">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="18941-800">- 内容</span><span class="sxs-lookup"><span data-stu-id="18941-800">- Content</span></span><br><span data-ttu-id="18941-801">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-801">
         - TaskPane</span></span><br><span data-ttu-id="18941-802">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-802">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="18941-804">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="18941-804">- ActiveView</span></span><br><span data-ttu-id="18941-805">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-805">
         - CompressedFile</span></span><br><span data-ttu-id="18941-806">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-806">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-807">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-807">
         - File</span></span><br><span data-ttu-id="18941-808">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-808">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-809">
         - PdfFile</span></span><br><span data-ttu-id="18941-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-810">
         - Selection</span></span><br><span data-ttu-id="18941-811">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-811">
         - Settings</span></span><br><span data-ttu-id="18941-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-813">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="18941-813">Office 2019 for Mac</span></span><br><span data-ttu-id="18941-814">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-814">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-815">- 内容</span><span class="sxs-lookup"><span data-stu-id="18941-815">- Content</span></span><br><span data-ttu-id="18941-816">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-816">
         - TaskPane</span></span><br><span data-ttu-id="18941-817">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-817">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-818">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-818">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="18941-819">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="18941-819">- ActiveView</span></span><br><span data-ttu-id="18941-820">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-820">
         - CompressedFile</span></span><br><span data-ttu-id="18941-821">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-821">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-822">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-822">
         - File</span></span><br><span data-ttu-id="18941-823">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-823">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-824">
         - PdfFile</span></span><br><span data-ttu-id="18941-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-825">
         - Selection</span></span><br><span data-ttu-id="18941-826">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-826">
         - Settings</span></span><br><span data-ttu-id="18941-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-828">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="18941-828">Office 2016 for Mac</span></span><br><span data-ttu-id="18941-829">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-829">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-830">- 内容</span><span class="sxs-lookup"><span data-stu-id="18941-830">- Content</span></span><br><span data-ttu-id="18941-831">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-831">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="18941-832">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="18941-832">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="18941-833">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="18941-833">- ActiveView</span></span><br><span data-ttu-id="18941-834">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="18941-834">
         - CompressedFile</span></span><br><span data-ttu-id="18941-835">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-835">
         - DocumentEvents</span></span><br><span data-ttu-id="18941-836">
         - File</span><span class="sxs-lookup"><span data-stu-id="18941-836">
         - File</span></span><br><span data-ttu-id="18941-837">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-837">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-838">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="18941-838">
         - PdfFile</span></span><br><span data-ttu-id="18941-839">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="18941-839">
         - Selection</span></span><br><span data-ttu-id="18941-840">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-840">
         - Settings</span></span><br><span data-ttu-id="18941-841">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-841">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="18941-842">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="18941-842">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="18941-843">OneNote</span><span class="sxs-lookup"><span data-stu-id="18941-843">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="18941-844">平台</span><span class="sxs-lookup"><span data-stu-id="18941-844">Platform</span></span></th>
    <th><span data-ttu-id="18941-845">扩展点</span><span class="sxs-lookup"><span data-stu-id="18941-845">Extension points</span></span></th>
    <th><span data-ttu-id="18941-846">API 要求集</span><span class="sxs-lookup"><span data-stu-id="18941-846">API requirement sets</span></span></th>
    <th><span data-ttu-id="18941-847"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="18941-847"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-848">Office Online</span><span class="sxs-lookup"><span data-stu-id="18941-848">Office Online</span></span></td>
    <td> <span data-ttu-id="18941-849">- 内容</span><span class="sxs-lookup"><span data-stu-id="18941-849">- Content</span></span><br><span data-ttu-id="18941-850">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-850">
         - TaskPane</span></span><br><span data-ttu-id="18941-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="18941-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="18941-852">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-852">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="18941-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="18941-854">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="18941-854">- DocumentEvents</span></span><br><span data-ttu-id="18941-855">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-855">
         - HtmlCoercion</span></span><br><span data-ttu-id="18941-856">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-856">
         - ImageCoercion</span></span><br><span data-ttu-id="18941-857">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="18941-857">
         - Settings</span></span><br><span data-ttu-id="18941-858">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-858">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="18941-859">项目</span><span class="sxs-lookup"><span data-stu-id="18941-859">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="18941-860">平台</span><span class="sxs-lookup"><span data-stu-id="18941-860">Platform</span></span></th>
    <th><span data-ttu-id="18941-861">扩展点</span><span class="sxs-lookup"><span data-stu-id="18941-861">Extension points</span></span></th>
    <th><span data-ttu-id="18941-862">API 要求集</span><span class="sxs-lookup"><span data-stu-id="18941-862">API requirement sets</span></span></th>
    <th><span data-ttu-id="18941-863"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="18941-863"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-864">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="18941-864">Office 2019 for Windows</span></span><br><span data-ttu-id="18941-865">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-865">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-866">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-866">- TaskPane</span></span></td>
    <td> <span data-ttu-id="18941-867">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-867">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="18941-868">- Selection</span><span class="sxs-lookup"><span data-stu-id="18941-868">- Selection</span></span><br><span data-ttu-id="18941-869">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-869">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-870">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="18941-870">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="18941-871">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-871">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-872">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-872">- TaskPane</span></span></td>
    <td> <span data-ttu-id="18941-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="18941-874">- Selection</span><span class="sxs-lookup"><span data-stu-id="18941-874">- Selection</span></span><br><span data-ttu-id="18941-875">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-875">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="18941-876">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="18941-876">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="18941-877">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="18941-877">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="18941-878">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="18941-878">- TaskPane</span></span></td>
    <td> <span data-ttu-id="18941-879">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="18941-879">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="18941-880">- Selection</span><span class="sxs-lookup"><span data-stu-id="18941-880">- Selection</span></span><br><span data-ttu-id="18941-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="18941-881">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="18941-882">另请参阅</span><span class="sxs-lookup"><span data-stu-id="18941-882">See also</span></span>

- [<span data-ttu-id="18941-883">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="18941-883">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="18941-884">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="18941-884">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="18941-885">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="18941-885">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="18941-886">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="18941-886">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="18941-887">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="18941-887">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="18941-888">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="18941-888">Update history for Office 365 ProPlus releases</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="18941-889">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="18941-889">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="18941-890">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="18941-890">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="18941-891">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="18941-891">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="18941-892">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="18941-892">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="18941-893">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="18941-893">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)

---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 04/13/2020
localization_priority: Priority
ms.openlocfilehash: 72da8db755fe6d1d166f66a70c8c298e5a27adff
ms.sourcegitcommit: 118e8bcbcfb73c93e2053bda67fe8dd20799b170
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/13/2020
ms.locfileid: "43241054"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="29dfb-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="29dfb-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="29dfb-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="29dfb-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="29dfb-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="29dfb-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="29dfb-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="29dfb-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="29dfb-108">Excel</span><span class="sxs-lookup"><span data-stu-id="29dfb-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="29dfb-109">平台</span><span class="sxs-lookup"><span data-stu-id="29dfb-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="29dfb-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="29dfb-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="29dfb-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="29dfb-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="29dfb-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="29dfb-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="29dfb-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="29dfb-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-114">- TaskPane</span></span><br><span data-ttu-id="29dfb-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-115">
        - Content</span></span><br><span data-ttu-id="29dfb-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="29dfb-116">
        - Custom Functions</span></span><br><span data-ttu-id="29dfb-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="29dfb-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="29dfb-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="29dfb-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="29dfb-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="29dfb-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="29dfb-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="29dfb-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="29dfb-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="29dfb-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="29dfb-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="29dfb-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="29dfb-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="29dfb-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="29dfb-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-130">
        - BindingEvents</span></span><br><span data-ttu-id="29dfb-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-131">
        - CompressedFile</span></span><br><span data-ttu-id="29dfb-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-132">
        - DocumentEvents</span></span><br><span data-ttu-id="29dfb-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-133">
        - File</span></span><br><span data-ttu-id="29dfb-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-134">
        - MatrixBindings</span></span><br><span data-ttu-id="29dfb-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-136">
        - Selection</span></span><br><span data-ttu-id="29dfb-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-137">
        - Settings</span></span><br><span data-ttu-id="29dfb-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-138">
        - TableBindings</span></span><br><span data-ttu-id="29dfb-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-139">
        - TableCoercion</span></span><br><span data-ttu-id="29dfb-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-140">
        - TextBindings</span></span><br><span data-ttu-id="29dfb-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-142">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="29dfb-142">Office on Windows</span></span><br><span data-ttu-id="29dfb-143">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="29dfb-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="29dfb-144">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-144">- TaskPane</span></span><br><span data-ttu-id="29dfb-145">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-145">
        - Content</span></span><br><span data-ttu-id="29dfb-146">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="29dfb-146">
        - Custom Functions</span></span><br><span data-ttu-id="29dfb-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="29dfb-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="29dfb-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="29dfb-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="29dfb-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="29dfb-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="29dfb-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="29dfb-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="29dfb-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="29dfb-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="29dfb-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="29dfb-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="29dfb-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="29dfb-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="29dfb-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-161">
        - BindingEvents</span></span><br><span data-ttu-id="29dfb-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-162">
        - CompressedFile</span></span><br><span data-ttu-id="29dfb-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-163">
        - DocumentEvents</span></span><br><span data-ttu-id="29dfb-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-164">
        - File</span></span><br><span data-ttu-id="29dfb-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-165">
        - MatrixBindings</span></span><br><span data-ttu-id="29dfb-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-167">
        - Selection</span></span><br><span data-ttu-id="29dfb-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-168">
        - Settings</span></span><br><span data-ttu-id="29dfb-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-169">
        - TableBindings</span></span><br><span data-ttu-id="29dfb-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-170">
        - TableCoercion</span></span><br><span data-ttu-id="29dfb-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-171">
        - TextBindings</span></span><br><span data-ttu-id="29dfb-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-173">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="29dfb-173">Office 2019 on Windows</span></span><br><span data-ttu-id="29dfb-174">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="29dfb-175">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-175">- TaskPane</span></span><br><span data-ttu-id="29dfb-176">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-176">
        - Content</span></span><br><span data-ttu-id="29dfb-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="29dfb-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="29dfb-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="29dfb-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="29dfb-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="29dfb-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="29dfb-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="29dfb-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="29dfb-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="29dfb-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="29dfb-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-188">- BindingEvents</span></span><br><span data-ttu-id="29dfb-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-189">
        - CompressedFile</span></span><br><span data-ttu-id="29dfb-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-190">
        - DocumentEvents</span></span><br><span data-ttu-id="29dfb-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-191">
        - File</span></span><br><span data-ttu-id="29dfb-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-192">
        - MatrixBindings</span></span><br><span data-ttu-id="29dfb-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-194">
        - Selection</span></span><br><span data-ttu-id="29dfb-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-195">
        - Settings</span></span><br><span data-ttu-id="29dfb-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-196">
        - TableBindings</span></span><br><span data-ttu-id="29dfb-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-197">
        - TableCoercion</span></span><br><span data-ttu-id="29dfb-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-198">
        - TextBindings</span></span><br><span data-ttu-id="29dfb-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-200">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="29dfb-200">Office 2016 on Windows</span></span><br><span data-ttu-id="29dfb-201">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="29dfb-202">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-202">- TaskPane</span></span><br><span data-ttu-id="29dfb-203">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-203">
        - Content</span></span></td>
    <td><span data-ttu-id="29dfb-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="29dfb-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="29dfb-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="29dfb-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="29dfb-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-207">- BindingEvents</span></span><br><span data-ttu-id="29dfb-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-208">
        - CompressedFile</span></span><br><span data-ttu-id="29dfb-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-209">
        - DocumentEvents</span></span><br><span data-ttu-id="29dfb-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-210">
        - File</span></span><br><span data-ttu-id="29dfb-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-211">
        - MatrixBindings</span></span><br><span data-ttu-id="29dfb-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-213">
        - Selection</span></span><br><span data-ttu-id="29dfb-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-214">
        - Settings</span></span><br><span data-ttu-id="29dfb-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-215">
        - TableBindings</span></span><br><span data-ttu-id="29dfb-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-216">
        - TableCoercion</span></span><br><span data-ttu-id="29dfb-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-217">
        - TextBindings</span></span><br><span data-ttu-id="29dfb-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-219">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="29dfb-219">Office 2013 on Windows</span></span><br><span data-ttu-id="29dfb-220">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="29dfb-221">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-221">
        - TaskPane</span></span><br><span data-ttu-id="29dfb-222">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="29dfb-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="29dfb-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="29dfb-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="29dfb-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-225">
        - BindingEvents</span></span><br><span data-ttu-id="29dfb-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-226">
        - CompressedFile</span></span><br><span data-ttu-id="29dfb-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-227">
        - DocumentEvents</span></span><br><span data-ttu-id="29dfb-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-228">
        - File</span></span><br><span data-ttu-id="29dfb-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-229">
        - MatrixBindings</span></span><br><span data-ttu-id="29dfb-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-231">
        - Selection</span></span><br><span data-ttu-id="29dfb-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-232">
        - Settings</span></span><br><span data-ttu-id="29dfb-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-233">
        - TableBindings</span></span><br><span data-ttu-id="29dfb-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-234">
        - TableCoercion</span></span><br><span data-ttu-id="29dfb-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-235">
        - TextBindings</span></span><br><span data-ttu-id="29dfb-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-237">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="29dfb-237">Office on iPad</span></span><br><span data-ttu-id="29dfb-238">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="29dfb-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="29dfb-239">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-239">- TaskPane</span></span><br><span data-ttu-id="29dfb-240">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-240">
        - Content</span></span></td>
    <td><span data-ttu-id="29dfb-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="29dfb-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="29dfb-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="29dfb-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="29dfb-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="29dfb-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="29dfb-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="29dfb-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="29dfb-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="29dfb-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="29dfb-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="29dfb-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-253">- BindingEvents</span></span><br><span data-ttu-id="29dfb-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-254">
        - DocumentEvents</span></span><br><span data-ttu-id="29dfb-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-255">
        - File</span></span><br><span data-ttu-id="29dfb-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-256">
        - MatrixBindings</span></span><br><span data-ttu-id="29dfb-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-258">
        - Selection</span></span><br><span data-ttu-id="29dfb-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-259">
        - Settings</span></span><br><span data-ttu-id="29dfb-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-260">
        - TableBindings</span></span><br><span data-ttu-id="29dfb-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-261">
        - TableCoercion</span></span><br><span data-ttu-id="29dfb-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-262">
        - TextBindings</span></span><br><span data-ttu-id="29dfb-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-264">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="29dfb-264">Office on Mac</span></span><br><span data-ttu-id="29dfb-265">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="29dfb-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="29dfb-266">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-266">- TaskPane</span></span><br><span data-ttu-id="29dfb-267">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-267">
        - Content</span></span><br><span data-ttu-id="29dfb-268">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="29dfb-268">
        - Custom Functions</span></span><br><span data-ttu-id="29dfb-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="29dfb-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="29dfb-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="29dfb-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="29dfb-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="29dfb-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="29dfb-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="29dfb-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="29dfb-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="29dfb-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="29dfb-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="29dfb-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="29dfb-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="29dfb-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-283">- BindingEvents</span></span><br><span data-ttu-id="29dfb-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-284">
        - CompressedFile</span></span><br><span data-ttu-id="29dfb-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-285">
        - DocumentEvents</span></span><br><span data-ttu-id="29dfb-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-286">
        - File</span></span><br><span data-ttu-id="29dfb-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-287">
        - MatrixBindings</span></span><br><span data-ttu-id="29dfb-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-289">
        - PdfFile</span></span><br><span data-ttu-id="29dfb-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-290">
        - Selection</span></span><br><span data-ttu-id="29dfb-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-291">
        - Settings</span></span><br><span data-ttu-id="29dfb-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-292">
        - TableBindings</span></span><br><span data-ttu-id="29dfb-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-293">
        - TableCoercion</span></span><br><span data-ttu-id="29dfb-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-294">
        - TextBindings</span></span><br><span data-ttu-id="29dfb-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-296">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="29dfb-296">Office 2019 on Mac</span></span><br><span data-ttu-id="29dfb-297">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="29dfb-298">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-298">- TaskPane</span></span><br><span data-ttu-id="29dfb-299">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-299">
        - Content</span></span><br><span data-ttu-id="29dfb-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="29dfb-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="29dfb-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="29dfb-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="29dfb-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="29dfb-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="29dfb-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="29dfb-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="29dfb-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="29dfb-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="29dfb-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-311">- BindingEvents</span></span><br><span data-ttu-id="29dfb-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-312">
        - CompressedFile</span></span><br><span data-ttu-id="29dfb-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-313">
        - DocumentEvents</span></span><br><span data-ttu-id="29dfb-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-314">
        - File</span></span><br><span data-ttu-id="29dfb-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-315">
        - MatrixBindings</span></span><br><span data-ttu-id="29dfb-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-317">
        - PdfFile</span></span><br><span data-ttu-id="29dfb-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-318">
        - Selection</span></span><br><span data-ttu-id="29dfb-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-319">
        - Settings</span></span><br><span data-ttu-id="29dfb-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-320">
        - TableBindings</span></span><br><span data-ttu-id="29dfb-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-321">
        - TableCoercion</span></span><br><span data-ttu-id="29dfb-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-322">
        - TextBindings</span></span><br><span data-ttu-id="29dfb-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-324">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="29dfb-324">Office 2016 on Mac</span></span><br><span data-ttu-id="29dfb-325">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="29dfb-326">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-326">- TaskPane</span></span><br><span data-ttu-id="29dfb-327">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-327">
        - Content</span></span></td>
    <td><span data-ttu-id="29dfb-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="29dfb-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="29dfb-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="29dfb-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="29dfb-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-331">- BindingEvents</span></span><br><span data-ttu-id="29dfb-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-332">
        - CompressedFile</span></span><br><span data-ttu-id="29dfb-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-333">
        - DocumentEvents</span></span><br><span data-ttu-id="29dfb-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-334">
        - File</span></span><br><span data-ttu-id="29dfb-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-335">
        - MatrixBindings</span></span><br><span data-ttu-id="29dfb-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-337">
        - PdfFile</span></span><br><span data-ttu-id="29dfb-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-338">
        - Selection</span></span><br><span data-ttu-id="29dfb-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-339">
        - Settings</span></span><br><span data-ttu-id="29dfb-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-340">
        - TableBindings</span></span><br><span data-ttu-id="29dfb-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-341">
        - TableCoercion</span></span><br><span data-ttu-id="29dfb-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-342">
        - TextBindings</span></span><br><span data-ttu-id="29dfb-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="29dfb-344">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="29dfb-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="29dfb-345">自定义函数（仅 Excel）</span><span class="sxs-lookup"><span data-stu-id="29dfb-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="29dfb-346">平台</span><span class="sxs-lookup"><span data-stu-id="29dfb-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="29dfb-347">扩展点</span><span class="sxs-lookup"><span data-stu-id="29dfb-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="29dfb-348">API 要求集</span><span class="sxs-lookup"><span data-stu-id="29dfb-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="29dfb-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="29dfb-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-350">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="29dfb-350">Office on the web</span></span></td>
    <td><span data-ttu-id="29dfb-351">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="29dfb-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="29dfb-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-353">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="29dfb-353">Office on Windows</span></span><br><span data-ttu-id="29dfb-354">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="29dfb-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="29dfb-355">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="29dfb-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="29dfb-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-357">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="29dfb-357">Office for Mac</span></span><br><span data-ttu-id="29dfb-358">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="29dfb-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="29dfb-359">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="29dfb-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="29dfb-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="29dfb-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="29dfb-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="29dfb-362">平台</span><span class="sxs-lookup"><span data-stu-id="29dfb-362">Platform</span></span></th>
    <th><span data-ttu-id="29dfb-363">扩展点</span><span class="sxs-lookup"><span data-stu-id="29dfb-363">Extension points</span></span></th>
    <th><span data-ttu-id="29dfb-364">API 要求集</span><span class="sxs-lookup"><span data-stu-id="29dfb-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="29dfb-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="29dfb-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-366">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="29dfb-366">Office on the web</span></span><br><span data-ttu-id="29dfb-367">（新式）</span><span class="sxs-lookup"><span data-stu-id="29dfb-367">(modern)</span></span></td>
    <td> <span data-ttu-id="29dfb-368">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-368">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="29dfb-369">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-369">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="29dfb-370">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-370">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="29dfb-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="29dfb-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="29dfb-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="29dfb-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="29dfb-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="29dfb-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="29dfb-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="29dfb-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="29dfb-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="29dfb-381">不可用</span><span class="sxs-lookup"><span data-stu-id="29dfb-381">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-382">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="29dfb-382">Office on the web</span></span><br><span data-ttu-id="29dfb-383">（经典）</span><span class="sxs-lookup"><span data-stu-id="29dfb-383">(classic)</span></span></td>
    <td> <span data-ttu-id="29dfb-384">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-384">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="29dfb-385">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-385">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="29dfb-386">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-386">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="29dfb-387">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-387">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="29dfb-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="29dfb-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="29dfb-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="29dfb-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="29dfb-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="29dfb-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="29dfb-395">不可用</span><span class="sxs-lookup"><span data-stu-id="29dfb-395">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-396">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="29dfb-396">Office on Windows</span></span><br><span data-ttu-id="29dfb-397">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="29dfb-397">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="29dfb-398">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-398">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="29dfb-399">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-399">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="29dfb-400">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-400">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="29dfb-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="29dfb-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="29dfb-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="29dfb-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="29dfb-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="29dfb-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="29dfb-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="29dfb-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="29dfb-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="29dfb-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="29dfb-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="29dfb-412">不可用</span><span class="sxs-lookup"><span data-stu-id="29dfb-412">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-413">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="29dfb-413">Office 2019 on Windows</span></span><br><span data-ttu-id="29dfb-414">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-414">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-415">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-415">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="29dfb-416">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-416">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="29dfb-417">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-417">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="29dfb-418">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-418">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="29dfb-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="29dfb-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="29dfb-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="29dfb-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="29dfb-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="29dfb-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="29dfb-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="29dfb-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="29dfb-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="29dfb-428">不可用</span><span class="sxs-lookup"><span data-stu-id="29dfb-428">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-429">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="29dfb-429">Office 2016 on Windows</span></span><br><span data-ttu-id="29dfb-430">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-430">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-431">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-431">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="29dfb-432">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-432">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="29dfb-433">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-433">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="29dfb-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="29dfb-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="29dfb-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="29dfb-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="29dfb-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="29dfb-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="29dfb-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="29dfb-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="29dfb-441">不可用</span><span class="sxs-lookup"><span data-stu-id="29dfb-441">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-442">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="29dfb-442">Office 2013 on Windows</span></span><br><span data-ttu-id="29dfb-443">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-443">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-444">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-444">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="29dfb-445">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-445">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="29dfb-446">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-446">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="29dfb-447">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-447">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="29dfb-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="29dfb-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="29dfb-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="29dfb-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="29dfb-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="29dfb-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="29dfb-452">不可用</span><span class="sxs-lookup"><span data-stu-id="29dfb-452">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-453">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="29dfb-453">Office on iOS</span></span><br><span data-ttu-id="29dfb-454">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="29dfb-454">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="29dfb-455">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-455">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="29dfb-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="29dfb-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="29dfb-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="29dfb-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="29dfb-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="29dfb-462">不可用</span><span class="sxs-lookup"><span data-stu-id="29dfb-462">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-463">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="29dfb-463">Office on Mac</span></span><br><span data-ttu-id="29dfb-464">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="29dfb-464">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="29dfb-465">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-465">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="29dfb-466">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-466">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="29dfb-467">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-467">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="29dfb-468">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-468">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="29dfb-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="29dfb-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="29dfb-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="29dfb-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="29dfb-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="29dfb-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="29dfb-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="29dfb-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="29dfb-478">不可用</span><span class="sxs-lookup"><span data-stu-id="29dfb-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-479">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="29dfb-479">Office 2019 on Mac</span></span><br><span data-ttu-id="29dfb-480">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-480">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-481">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-481">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="29dfb-482">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-482">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="29dfb-483">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-483">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="29dfb-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="29dfb-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="29dfb-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="29dfb-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="29dfb-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="29dfb-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="29dfb-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="29dfb-492">不可用</span><span class="sxs-lookup"><span data-stu-id="29dfb-492">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-493">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="29dfb-493">Office 2016 on Mac</span></span><br><span data-ttu-id="29dfb-494">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-494">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-495">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-495">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="29dfb-496">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-496">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="29dfb-497">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-497">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="29dfb-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="29dfb-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="29dfb-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="29dfb-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="29dfb-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="29dfb-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="29dfb-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="29dfb-506">不可用</span><span class="sxs-lookup"><span data-stu-id="29dfb-506">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-507">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="29dfb-507">Office on Android</span></span><br><span data-ttu-id="29dfb-508">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="29dfb-508">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="29dfb-509">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-509">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="29dfb-510">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">约会组织者（撰写）：联机会议</a> （预览）</span><span class="sxs-lookup"><span data-stu-id="29dfb-510">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="29dfb-511">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-511">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-512">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-512">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="29dfb-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="29dfb-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="29dfb-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="29dfb-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="29dfb-517">不可用</span><span class="sxs-lookup"><span data-stu-id="29dfb-517">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="29dfb-518">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="29dfb-518">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="29dfb-519">要求集的客户端支持可能受到 Exchange 服务器支持的限制。</span><span class="sxs-lookup"><span data-stu-id="29dfb-519">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="29dfb-520">有关 Exchange 服务器和 Outlook 客户端支持的要求集范围的详细信息，请参阅 [Outlook JavaScript API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="29dfb-520">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="29dfb-521">Word</span><span class="sxs-lookup"><span data-stu-id="29dfb-521">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="29dfb-522">平台</span><span class="sxs-lookup"><span data-stu-id="29dfb-522">Platform</span></span></th>
    <th><span data-ttu-id="29dfb-523">扩展点</span><span class="sxs-lookup"><span data-stu-id="29dfb-523">Extension points</span></span></th>
    <th><span data-ttu-id="29dfb-524">API 要求集</span><span class="sxs-lookup"><span data-stu-id="29dfb-524">API requirement sets</span></span></th>
    <th><span data-ttu-id="29dfb-525"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="29dfb-525"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-526">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="29dfb-526">Office on the web</span></span></td>
    <td> <span data-ttu-id="29dfb-527">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-527">- TaskPane</span></span><br><span data-ttu-id="29dfb-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="29dfb-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="29dfb-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="29dfb-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="29dfb-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="29dfb-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-535">- BindingEvents</span></span><br><span data-ttu-id="29dfb-536">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="29dfb-536">
         - CustomXmlParts</span></span><br><span data-ttu-id="29dfb-537">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-537">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-538">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-538">
         - File</span></span><br><span data-ttu-id="29dfb-539">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-539">
         - HtmlCoercion</span></span><br><span data-ttu-id="29dfb-540">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-540">
         - MatrixBindings</span></span><br><span data-ttu-id="29dfb-541">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-541">
         - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-542">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-542">
         - OoxmlCoercion</span></span><br><span data-ttu-id="29dfb-543">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-543">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-544">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-544">
         - Selection</span></span><br><span data-ttu-id="29dfb-545">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-545">
         - Settings</span></span><br><span data-ttu-id="29dfb-546">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-546">
         - TableBindings</span></span><br><span data-ttu-id="29dfb-547">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-547">
         - TableCoercion</span></span><br><span data-ttu-id="29dfb-548">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-548">
         - TextBindings</span></span><br><span data-ttu-id="29dfb-549">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-549">
         - TextCoercion</span></span><br><span data-ttu-id="29dfb-550">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-550">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-551">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="29dfb-551">Office on Windows</span></span><br><span data-ttu-id="29dfb-552">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="29dfb-552">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="29dfb-553">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-553">- TaskPane</span></span><br><span data-ttu-id="29dfb-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="29dfb-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="29dfb-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="29dfb-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="29dfb-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="29dfb-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-561">- BindingEvents</span></span><br><span data-ttu-id="29dfb-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-562">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="29dfb-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="29dfb-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-564">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-565">
         - File</span></span><br><span data-ttu-id="29dfb-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="29dfb-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-567">
         - MatrixBindings</span></span><br><span data-ttu-id="29dfb-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="29dfb-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-570">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-571">
         - Selection</span></span><br><span data-ttu-id="29dfb-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-572">
         - Settings</span></span><br><span data-ttu-id="29dfb-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-573">
         - TableBindings</span></span><br><span data-ttu-id="29dfb-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-574">
         - TableCoercion</span></span><br><span data-ttu-id="29dfb-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-575">
         - TextBindings</span></span><br><span data-ttu-id="29dfb-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-576">
         - TextCoercion</span></span><br><span data-ttu-id="29dfb-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-578">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="29dfb-578">Office 2019 on Windows</span></span><br><span data-ttu-id="29dfb-579">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-580">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-580">- TaskPane</span></span><br><span data-ttu-id="29dfb-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-582">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-582">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="29dfb-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="29dfb-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="29dfb-585">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-585">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-586">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-586">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="29dfb-587">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-587">- BindingEvents</span></span><br><span data-ttu-id="29dfb-588">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-588">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-589">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="29dfb-589">
         - CustomXmlParts</span></span><br><span data-ttu-id="29dfb-590">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-590">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-591">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-591">
         - File</span></span><br><span data-ttu-id="29dfb-592">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-592">
         - HtmlCoercion</span></span><br><span data-ttu-id="29dfb-593">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-593">
         - MatrixBindings</span></span><br><span data-ttu-id="29dfb-594">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-594">
         - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-595">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-595">
         - OoxmlCoercion</span></span><br><span data-ttu-id="29dfb-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-596">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-597">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-597">
         - Selection</span></span><br><span data-ttu-id="29dfb-598">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-598">
         - Settings</span></span><br><span data-ttu-id="29dfb-599">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-599">
         - TableBindings</span></span><br><span data-ttu-id="29dfb-600">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-600">
         - TableCoercion</span></span><br><span data-ttu-id="29dfb-601">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-601">
         - TextBindings</span></span><br><span data-ttu-id="29dfb-602">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-602">
         - TextCoercion</span></span><br><span data-ttu-id="29dfb-603">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-603">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-604">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="29dfb-604">Office 2016 on Windows</span></span><br><span data-ttu-id="29dfb-605">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-605">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-606">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-606">- TaskPane</span></span></td>
    <td> <span data-ttu-id="29dfb-607">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-607">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="29dfb-608">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="29dfb-608">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="29dfb-609">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-609">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="29dfb-610">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-610">- BindingEvents</span></span><br><span data-ttu-id="29dfb-611">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-611">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-612">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="29dfb-612">
         - CustomXmlParts</span></span><br><span data-ttu-id="29dfb-613">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-613">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-614">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-614">
         - File</span></span><br><span data-ttu-id="29dfb-615">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-615">
         - HtmlCoercion</span></span><br><span data-ttu-id="29dfb-616">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-616">
         - MatrixBindings</span></span><br><span data-ttu-id="29dfb-617">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-617">
         - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-618">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-618">
         - OoxmlCoercion</span></span><br><span data-ttu-id="29dfb-619">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-619">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-620">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-620">
         - Selection</span></span><br><span data-ttu-id="29dfb-621">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-621">
         - Settings</span></span><br><span data-ttu-id="29dfb-622">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-622">
         - TableBindings</span></span><br><span data-ttu-id="29dfb-623">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-623">
         - TableCoercion</span></span><br><span data-ttu-id="29dfb-624">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-624">
         - TextBindings</span></span><br><span data-ttu-id="29dfb-625">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-625">
         - TextCoercion</span></span><br><span data-ttu-id="29dfb-626">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-626">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-627">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="29dfb-627">Office 2013 on Windows</span></span><br><span data-ttu-id="29dfb-628">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-628">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-629">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-629">- TaskPane</span></span></td>
    <td> <span data-ttu-id="29dfb-630">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="29dfb-630">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="29dfb-631">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-631">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="29dfb-632">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-632">- BindingEvents</span></span><br><span data-ttu-id="29dfb-633">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-633">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-634">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="29dfb-634">
         - CustomXmlParts</span></span><br><span data-ttu-id="29dfb-635">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-635">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-636">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-636">
         - File</span></span><br><span data-ttu-id="29dfb-637">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-637">
         - HtmlCoercion</span></span><br><span data-ttu-id="29dfb-638">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-638">
         - MatrixBindings</span></span><br><span data-ttu-id="29dfb-639">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-639">
         - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-640">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-640">
         - OoxmlCoercion</span></span><br><span data-ttu-id="29dfb-641">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-641">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-642">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-642">
         - Selection</span></span><br><span data-ttu-id="29dfb-643">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-643">
         - Settings</span></span><br><span data-ttu-id="29dfb-644">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-644">
         - TableBindings</span></span><br><span data-ttu-id="29dfb-645">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-645">
         - TableCoercion</span></span><br><span data-ttu-id="29dfb-646">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-646">
         - TextBindings</span></span><br><span data-ttu-id="29dfb-647">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-647">
         - TextCoercion</span></span><br><span data-ttu-id="29dfb-648">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-648">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-649">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="29dfb-649">Office on iPad</span></span><br><span data-ttu-id="29dfb-650">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="29dfb-650">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="29dfb-651">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-651">- TaskPane</span></span></td>
    <td> <span data-ttu-id="29dfb-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="29dfb-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="29dfb-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="29dfb-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="29dfb-657">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-657">- BindingEvents</span></span><br><span data-ttu-id="29dfb-658">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-658">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-659">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="29dfb-659">
         - CustomXmlParts</span></span><br><span data-ttu-id="29dfb-660">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-660">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-661">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-661">
         - File</span></span><br><span data-ttu-id="29dfb-662">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-662">
         - HtmlCoercion</span></span><br><span data-ttu-id="29dfb-663">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-663">
         - MatrixBindings</span></span><br><span data-ttu-id="29dfb-664">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-664">
         - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-665">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-665">
         - OoxmlCoercion</span></span><br><span data-ttu-id="29dfb-666">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-666">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-667">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-667">
         - Selection</span></span><br><span data-ttu-id="29dfb-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-668">
         - Settings</span></span><br><span data-ttu-id="29dfb-669">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-669">
         - TableBindings</span></span><br><span data-ttu-id="29dfb-670">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-670">
         - TableCoercion</span></span><br><span data-ttu-id="29dfb-671">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-671">
         - TextBindings</span></span><br><span data-ttu-id="29dfb-672">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-672">
         - TextCoercion</span></span><br><span data-ttu-id="29dfb-673">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-673">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-674">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="29dfb-674">Office on Mac</span></span><br><span data-ttu-id="29dfb-675">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="29dfb-675">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="29dfb-676">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-676">- TaskPane</span></span><br><span data-ttu-id="29dfb-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="29dfb-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="29dfb-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="29dfb-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="29dfb-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="29dfb-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-684">- BindingEvents</span></span><br><span data-ttu-id="29dfb-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-685">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="29dfb-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="29dfb-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-687">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-688">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-688">
         - File</span></span><br><span data-ttu-id="29dfb-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="29dfb-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-690">
         - MatrixBindings</span></span><br><span data-ttu-id="29dfb-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="29dfb-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-693">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-694">
         - Selection</span></span><br><span data-ttu-id="29dfb-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-695">
         - Settings</span></span><br><span data-ttu-id="29dfb-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-696">
         - TableBindings</span></span><br><span data-ttu-id="29dfb-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-697">
         - TableCoercion</span></span><br><span data-ttu-id="29dfb-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-698">
         - TextBindings</span></span><br><span data-ttu-id="29dfb-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-699">
         - TextCoercion</span></span><br><span data-ttu-id="29dfb-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-701">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="29dfb-701">Office 2019 on Mac</span></span><br><span data-ttu-id="29dfb-702">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-703">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-703">- TaskPane</span></span><br><span data-ttu-id="29dfb-704">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-704">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-705">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-705">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="29dfb-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="29dfb-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="29dfb-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="29dfb-710">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-710">- BindingEvents</span></span><br><span data-ttu-id="29dfb-711">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-711">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-712">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="29dfb-712">
         - CustomXmlParts</span></span><br><span data-ttu-id="29dfb-713">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-713">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-714">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-714">
         - File</span></span><br><span data-ttu-id="29dfb-715">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-715">
         - HtmlCoercion</span></span><br><span data-ttu-id="29dfb-716">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-716">
         - MatrixBindings</span></span><br><span data-ttu-id="29dfb-717">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-717">
         - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-718">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-718">
         - OoxmlCoercion</span></span><br><span data-ttu-id="29dfb-719">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-719">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-720">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-720">
         - Selection</span></span><br><span data-ttu-id="29dfb-721">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-721">
         - Settings</span></span><br><span data-ttu-id="29dfb-722">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-722">
         - TableBindings</span></span><br><span data-ttu-id="29dfb-723">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-723">
         - TableCoercion</span></span><br><span data-ttu-id="29dfb-724">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-724">
         - TextBindings</span></span><br><span data-ttu-id="29dfb-725">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-725">
         - TextCoercion</span></span><br><span data-ttu-id="29dfb-726">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-726">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-727">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="29dfb-727">Office 2016 on Mac</span></span><br><span data-ttu-id="29dfb-728">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-728">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-729">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-729">- TaskPane</span></span></td>
    <td> <span data-ttu-id="29dfb-730">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-730">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="29dfb-731">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="29dfb-731">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="29dfb-732">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-732">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="29dfb-733">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-733">- BindingEvents</span></span><br><span data-ttu-id="29dfb-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-734">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-735">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="29dfb-735">
         - CustomXmlParts</span></span><br><span data-ttu-id="29dfb-736">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-736">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-737">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-737">
         - File</span></span><br><span data-ttu-id="29dfb-738">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-738">
         - HtmlCoercion</span></span><br><span data-ttu-id="29dfb-739">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-739">
         - MatrixBindings</span></span><br><span data-ttu-id="29dfb-740">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-740">
         - MatrixCoercion</span></span><br><span data-ttu-id="29dfb-741">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-741">
         - OoxmlCoercion</span></span><br><span data-ttu-id="29dfb-742">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-742">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-743">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-743">
         - Selection</span></span><br><span data-ttu-id="29dfb-744">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-744">
         - Settings</span></span><br><span data-ttu-id="29dfb-745">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-745">
         - TableBindings</span></span><br><span data-ttu-id="29dfb-746">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-746">
         - TableCoercion</span></span><br><span data-ttu-id="29dfb-747">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="29dfb-747">
         - TextBindings</span></span><br><span data-ttu-id="29dfb-748">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-748">
         - TextCoercion</span></span><br><span data-ttu-id="29dfb-749">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-749">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="29dfb-750">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="29dfb-750">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="29dfb-751">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="29dfb-751">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="29dfb-752">平台</span><span class="sxs-lookup"><span data-stu-id="29dfb-752">Platform</span></span></th>
    <th><span data-ttu-id="29dfb-753">扩展点</span><span class="sxs-lookup"><span data-stu-id="29dfb-753">Extension points</span></span></th>
    <th><span data-ttu-id="29dfb-754">API 要求集</span><span class="sxs-lookup"><span data-stu-id="29dfb-754">API requirement sets</span></span></th>
    <th><span data-ttu-id="29dfb-755"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="29dfb-755"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-756">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="29dfb-756">Office on the web</span></span></td>
    <td> <span data-ttu-id="29dfb-757">- 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-757">- Content</span></span><br><span data-ttu-id="29dfb-758">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-758">
         - TaskPane</span></span><br><span data-ttu-id="29dfb-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-760">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-760">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="29dfb-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="29dfb-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="29dfb-764">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="29dfb-764">- ActiveView</span></span><br><span data-ttu-id="29dfb-765">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-765">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-766">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-766">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-767">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-767">
         - File</span></span><br><span data-ttu-id="29dfb-768">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-768">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-769">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-769">
         - Selection</span></span><br><span data-ttu-id="29dfb-770">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-770">
         - Settings</span></span><br><span data-ttu-id="29dfb-771">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-771">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-772">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="29dfb-772">Office on Windows</span></span><br><span data-ttu-id="29dfb-773">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="29dfb-773">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="29dfb-774">- 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-774">- Content</span></span><br><span data-ttu-id="29dfb-775">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-775">
         - TaskPane</span></span><br><span data-ttu-id="29dfb-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-777">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-777">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="29dfb-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="29dfb-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="29dfb-781">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="29dfb-781">- ActiveView</span></span><br><span data-ttu-id="29dfb-782">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-782">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-783">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-783">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-784">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-784">
         - File</span></span><br><span data-ttu-id="29dfb-785">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-785">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-786">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-786">
         - Selection</span></span><br><span data-ttu-id="29dfb-787">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-787">
         - Settings</span></span><br><span data-ttu-id="29dfb-788">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-788">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-789">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="29dfb-789">Office 2019 on Windows</span></span><br><span data-ttu-id="29dfb-790">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-790">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-791">- 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-791">- Content</span></span><br><span data-ttu-id="29dfb-792">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-792">
         - TaskPane</span></span><br><span data-ttu-id="29dfb-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="29dfb-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="29dfb-796">- ActiveView</span></span><br><span data-ttu-id="29dfb-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-797">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-798">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-799">
         - File</span></span><br><span data-ttu-id="29dfb-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-800">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-801">
         - Selection</span></span><br><span data-ttu-id="29dfb-802">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-802">
         - Settings</span></span><br><span data-ttu-id="29dfb-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-804">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="29dfb-804">Office 2016 on Windows</span></span><br><span data-ttu-id="29dfb-805">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-805">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-806">- 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-806">- Content</span></span><br><span data-ttu-id="29dfb-807">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="29dfb-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="29dfb-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="29dfb-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="29dfb-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="29dfb-810">- ActiveView</span></span><br><span data-ttu-id="29dfb-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-811">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-812">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-813">
         - File</span></span><br><span data-ttu-id="29dfb-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-814">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-815">
         - Selection</span></span><br><span data-ttu-id="29dfb-816">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-816">
         - Settings</span></span><br><span data-ttu-id="29dfb-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-818">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="29dfb-818">Office 2013 on Windows</span></span><br><span data-ttu-id="29dfb-819">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-819">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-820">- 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-820">- Content</span></span><br><span data-ttu-id="29dfb-821">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-821">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="29dfb-822">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="29dfb-822">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="29dfb-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="29dfb-824">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="29dfb-824">- ActiveView</span></span><br><span data-ttu-id="29dfb-825">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-825">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-826">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-826">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-827">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-827">
         - File</span></span><br><span data-ttu-id="29dfb-828">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-828">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-829">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-829">
         - Selection</span></span><br><span data-ttu-id="29dfb-830">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-830">
         - Settings</span></span><br><span data-ttu-id="29dfb-831">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-831">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-832">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="29dfb-832">Office on iPad</span></span><br><span data-ttu-id="29dfb-833">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="29dfb-833">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="29dfb-834">- 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-834">- Content</span></span><br><span data-ttu-id="29dfb-835">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-835">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="29dfb-836">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-836">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="29dfb-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="29dfb-839">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="29dfb-839">- ActiveView</span></span><br><span data-ttu-id="29dfb-840">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-840">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-841">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-841">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-842">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-842">
         - File</span></span><br><span data-ttu-id="29dfb-843">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-843">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-844">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-844">
         - Selection</span></span><br><span data-ttu-id="29dfb-845">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-845">
         - Settings</span></span><br><span data-ttu-id="29dfb-846">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-846">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-847">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="29dfb-847">Office on Mac</span></span><br><span data-ttu-id="29dfb-848">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="29dfb-848">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="29dfb-849">- 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-849">- Content</span></span><br><span data-ttu-id="29dfb-850">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-850">
         - TaskPane</span></span><br><span data-ttu-id="29dfb-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-852">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-852">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="29dfb-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="29dfb-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="29dfb-856">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="29dfb-856">- ActiveView</span></span><br><span data-ttu-id="29dfb-857">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-857">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-858">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-858">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-859">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-859">
         - File</span></span><br><span data-ttu-id="29dfb-860">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-860">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-861">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-861">
         - Selection</span></span><br><span data-ttu-id="29dfb-862">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-862">
         - Settings</span></span><br><span data-ttu-id="29dfb-863">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-863">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-864">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="29dfb-864">Office 2019 on Mac</span></span><br><span data-ttu-id="29dfb-865">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-865">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-866">- 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-866">- Content</span></span><br><span data-ttu-id="29dfb-867">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-867">
         - TaskPane</span></span><br><span data-ttu-id="29dfb-868">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-868">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-869">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-869">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-870">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-870">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="29dfb-871">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="29dfb-871">- ActiveView</span></span><br><span data-ttu-id="29dfb-872">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-872">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-873">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-873">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-874">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-874">
         - File</span></span><br><span data-ttu-id="29dfb-875">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-875">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-876">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-876">
         - Selection</span></span><br><span data-ttu-id="29dfb-877">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-877">
         - Settings</span></span><br><span data-ttu-id="29dfb-878">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-878">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-879">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="29dfb-879">Office 2016 on Mac</span></span><br><span data-ttu-id="29dfb-880">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-880">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-881">- 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-881">- Content</span></span><br><span data-ttu-id="29dfb-882">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-882">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="29dfb-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="29dfb-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="29dfb-884">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-884">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="29dfb-885">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="29dfb-885">- ActiveView</span></span><br><span data-ttu-id="29dfb-886">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-886">
         - CompressedFile</span></span><br><span data-ttu-id="29dfb-887">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-887">
         - DocumentEvents</span></span><br><span data-ttu-id="29dfb-888">
         - File</span><span class="sxs-lookup"><span data-stu-id="29dfb-888">
         - File</span></span><br><span data-ttu-id="29dfb-889">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="29dfb-889">
         - PdfFile</span></span><br><span data-ttu-id="29dfb-890">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-890">
         - Selection</span></span><br><span data-ttu-id="29dfb-891">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-891">
         - Settings</span></span><br><span data-ttu-id="29dfb-892">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-892">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="29dfb-893">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="29dfb-893">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="29dfb-894">OneNote</span><span class="sxs-lookup"><span data-stu-id="29dfb-894">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="29dfb-895">平台</span><span class="sxs-lookup"><span data-stu-id="29dfb-895">Platform</span></span></th>
    <th><span data-ttu-id="29dfb-896">扩展点</span><span class="sxs-lookup"><span data-stu-id="29dfb-896">Extension points</span></span></th>
    <th><span data-ttu-id="29dfb-897">API 要求集</span><span class="sxs-lookup"><span data-stu-id="29dfb-897">API requirement sets</span></span></th>
    <th><span data-ttu-id="29dfb-898"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="29dfb-898"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-899">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="29dfb-899">Office on the web</span></span></td>
    <td> <span data-ttu-id="29dfb-900">- 内容</span><span class="sxs-lookup"><span data-stu-id="29dfb-900">- Content</span></span><br><span data-ttu-id="29dfb-901">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-901">
         - TaskPane</span></span><br><span data-ttu-id="29dfb-902">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-902">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="29dfb-903">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-903">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="29dfb-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="29dfb-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="29dfb-906">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="29dfb-906">- DocumentEvents</span></span><br><span data-ttu-id="29dfb-907">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-907">
         - HtmlCoercion</span></span><br><span data-ttu-id="29dfb-908">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="29dfb-908">
         - Settings</span></span><br><span data-ttu-id="29dfb-909">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-909">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="29dfb-910">项目</span><span class="sxs-lookup"><span data-stu-id="29dfb-910">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="29dfb-911">平台</span><span class="sxs-lookup"><span data-stu-id="29dfb-911">Platform</span></span></th>
    <th><span data-ttu-id="29dfb-912">扩展点</span><span class="sxs-lookup"><span data-stu-id="29dfb-912">Extension points</span></span></th>
    <th><span data-ttu-id="29dfb-913">API 要求集</span><span class="sxs-lookup"><span data-stu-id="29dfb-913">API requirement sets</span></span></th>
    <th><span data-ttu-id="29dfb-914"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="29dfb-914"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-915">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="29dfb-915">Office 2019 on Windows</span></span><br><span data-ttu-id="29dfb-916">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-916">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-917">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-917">- TaskPane</span></span></td>
    <td> <span data-ttu-id="29dfb-918">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-918">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="29dfb-919">- Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-919">- Selection</span></span><br><span data-ttu-id="29dfb-920">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-920">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-921">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="29dfb-921">Office 2016 on Windows</span></span><br><span data-ttu-id="29dfb-922">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-922">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-923">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-923">- TaskPane</span></span></td>
    <td> <span data-ttu-id="29dfb-924">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-924">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="29dfb-925">- Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-925">- Selection</span></span><br><span data-ttu-id="29dfb-926">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-926">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="29dfb-927">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="29dfb-927">Office 2013 on Windows</span></span><br><span data-ttu-id="29dfb-928">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="29dfb-928">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="29dfb-929">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="29dfb-929">- TaskPane</span></span></td>
    <td> <span data-ttu-id="29dfb-930">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="29dfb-930">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="29dfb-931">- Selection</span><span class="sxs-lookup"><span data-stu-id="29dfb-931">- Selection</span></span><br><span data-ttu-id="29dfb-932">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="29dfb-932">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="29dfb-933">另请参阅</span><span class="sxs-lookup"><span data-stu-id="29dfb-933">See also</span></span>

- [<span data-ttu-id="29dfb-934">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="29dfb-934">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="29dfb-935">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="29dfb-935">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="29dfb-936">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="29dfb-936">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="29dfb-937">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="29dfb-937">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="29dfb-938">API 参考文档</span><span class="sxs-lookup"><span data-stu-id="29dfb-938">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="29dfb-939">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="29dfb-939">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="29dfb-940">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="29dfb-940">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="29dfb-941">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="29dfb-941">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="29dfb-942">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="29dfb-942">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="29dfb-943">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="29dfb-943">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="29dfb-944">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="29dfb-944">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="29dfb-945">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="29dfb-945">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
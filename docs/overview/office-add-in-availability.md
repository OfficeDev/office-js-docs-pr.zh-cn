---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 08/13/2019
localization_priority: Priority
ms.openlocfilehash: a3c580f32ad7cd384309a9b53e55ea488a470a90
ms.sourcegitcommit: f781d7cfd980cd866d6d1d00c5b9d16c8a4b7f9b
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/20/2019
ms.locfileid: "37053324"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="56df0-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="56df0-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="56df0-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="56df0-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="56df0-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="56df0-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="56df0-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="56df0-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="56df0-108">Excel</span><span class="sxs-lookup"><span data-stu-id="56df0-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="56df0-109">平台</span><span class="sxs-lookup"><span data-stu-id="56df0-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="56df0-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="56df0-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="56df0-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="56df0-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="56df0-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="56df0-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="56df0-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="56df0-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-114">- TaskPane</span></span><br><span data-ttu-id="56df0-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-115">
        - Content</span></span><br><span data-ttu-id="56df0-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="56df0-116">
        - Custom Functions</span></span><br><span data-ttu-id="56df0-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="56df0-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="56df0-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="56df0-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="56df0-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="56df0-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="56df0-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="56df0-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="56df0-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="56df0-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="56df0-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="56df0-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="56df0-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="56df0-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="56df0-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="56df0-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="56df0-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="56df0-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="56df0-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-128">
        - BindingEvents</span></span><br><span data-ttu-id="56df0-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-129">
        - CompressedFile</span></span><br><span data-ttu-id="56df0-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-130">
        - DocumentEvents</span></span><br><span data-ttu-id="56df0-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="56df0-131">
        - File</span></span><br><span data-ttu-id="56df0-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-132">
        - MatrixBindings</span></span><br><span data-ttu-id="56df0-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="56df0-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-134">
        - Selection</span></span><br><span data-ttu-id="56df0-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-135">
        - Settings</span></span><br><span data-ttu-id="56df0-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-136">
        - TableBindings</span></span><br><span data-ttu-id="56df0-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-137">
        - TableCoercion</span></span><br><span data-ttu-id="56df0-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-138">
        - TextBindings</span></span><br><span data-ttu-id="56df0-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-140">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="56df0-140">Office on Windows</span></span><br><span data-ttu-id="56df0-141">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="56df0-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="56df0-142">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-142">- TaskPane</span></span><br><span data-ttu-id="56df0-143">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-143">
        - Content</span></span><br><span data-ttu-id="56df0-144">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="56df0-144">
        - Custom Functions</span></span><br><span data-ttu-id="56df0-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="56df0-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="56df0-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="56df0-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="56df0-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="56df0-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="56df0-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="56df0-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="56df0-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="56df0-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="56df0-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="56df0-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="56df0-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="56df0-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="56df0-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="56df0-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="56df0-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="56df0-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="56df0-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="56df0-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-158">
        - BindingEvents</span></span><br><span data-ttu-id="56df0-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-159">
        - CompressedFile</span></span><br><span data-ttu-id="56df0-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-160">
        - DocumentEvents</span></span><br><span data-ttu-id="56df0-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="56df0-161">
        - File</span></span><br><span data-ttu-id="56df0-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-162">
        - MatrixBindings</span></span><br><span data-ttu-id="56df0-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="56df0-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-164">
        - Selection</span></span><br><span data-ttu-id="56df0-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-165">
        - Settings</span></span><br><span data-ttu-id="56df0-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-166">
        - TableBindings</span></span><br><span data-ttu-id="56df0-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-167">
        - TableCoercion</span></span><br><span data-ttu-id="56df0-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-168">
        - TextBindings</span></span><br><span data-ttu-id="56df0-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-170">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="56df0-170">Office 2019 on Windows</span></span><br><span data-ttu-id="56df0-171">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="56df0-172">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-172">- TaskPane</span></span><br><span data-ttu-id="56df0-173">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-173">
        - Content</span></span><br><span data-ttu-id="56df0-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="56df0-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="56df0-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="56df0-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="56df0-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="56df0-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="56df0-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="56df0-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="56df0-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="56df0-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="56df0-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="56df0-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="56df0-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="56df0-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="56df0-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="56df0-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-185">- BindingEvents</span></span><br><span data-ttu-id="56df0-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-186">
        - CompressedFile</span></span><br><span data-ttu-id="56df0-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-187">
        - DocumentEvents</span></span><br><span data-ttu-id="56df0-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="56df0-188">
        - File</span></span><br><span data-ttu-id="56df0-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-189">
        - MatrixBindings</span></span><br><span data-ttu-id="56df0-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="56df0-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-191">
        - Selection</span></span><br><span data-ttu-id="56df0-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-192">
        - Settings</span></span><br><span data-ttu-id="56df0-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-193">
        - TableBindings</span></span><br><span data-ttu-id="56df0-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-194">
        - TableCoercion</span></span><br><span data-ttu-id="56df0-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-195">
        - TextBindings</span></span><br><span data-ttu-id="56df0-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-197">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="56df0-197">Office 2016 on Windows</span></span><br><span data-ttu-id="56df0-198">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="56df0-199">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-199">- TaskPane</span></span><br><span data-ttu-id="56df0-200">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-200">
        - Content</span></span></td>
    <td><span data-ttu-id="56df0-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="56df0-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="56df0-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="56df0-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="56df0-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-204">- BindingEvents</span></span><br><span data-ttu-id="56df0-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-205">
        - CompressedFile</span></span><br><span data-ttu-id="56df0-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-206">
        - DocumentEvents</span></span><br><span data-ttu-id="56df0-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="56df0-207">
        - File</span></span><br><span data-ttu-id="56df0-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-208">
        - MatrixBindings</span></span><br><span data-ttu-id="56df0-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="56df0-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-210">
        - Selection</span></span><br><span data-ttu-id="56df0-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-211">
        - Settings</span></span><br><span data-ttu-id="56df0-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-212">
        - TableBindings</span></span><br><span data-ttu-id="56df0-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-213">
        - TableCoercion</span></span><br><span data-ttu-id="56df0-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-214">
        - TextBindings</span></span><br><span data-ttu-id="56df0-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-216">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="56df0-216">Office 2013 on Windows</span></span><br><span data-ttu-id="56df0-217">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="56df0-218">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-218">
        - TaskPane</span></span><br><span data-ttu-id="56df0-219">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="56df0-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="56df0-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="56df0-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="56df0-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-222">
        - BindingEvents</span></span><br><span data-ttu-id="56df0-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-223">
        - CompressedFile</span></span><br><span data-ttu-id="56df0-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-224">
        - DocumentEvents</span></span><br><span data-ttu-id="56df0-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="56df0-225">
        - File</span></span><br><span data-ttu-id="56df0-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-226">
        - MatrixBindings</span></span><br><span data-ttu-id="56df0-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="56df0-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-228">
        - Selection</span></span><br><span data-ttu-id="56df0-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-229">
        - Settings</span></span><br><span data-ttu-id="56df0-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-230">
        - TableBindings</span></span><br><span data-ttu-id="56df0-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-231">
        - TableCoercion</span></span><br><span data-ttu-id="56df0-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-232">
        - TextBindings</span></span><br><span data-ttu-id="56df0-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-234">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="56df0-234">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="56df0-235">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="56df0-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="56df0-236">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-236">- TaskPane</span></span><br><span data-ttu-id="56df0-237">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-237">
        - Content</span></span></td>
    <td><span data-ttu-id="56df0-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="56df0-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="56df0-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="56df0-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="56df0-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="56df0-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="56df0-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="56df0-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="56df0-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="56df0-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="56df0-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="56df0-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="56df0-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="56df0-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="56df0-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="56df0-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="56df0-249">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-249">- BindingEvents</span></span><br><span data-ttu-id="56df0-250">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-250">
        - DocumentEvents</span></span><br><span data-ttu-id="56df0-251">
        - File</span><span class="sxs-lookup"><span data-stu-id="56df0-251">
        - File</span></span><br><span data-ttu-id="56df0-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-252">
        - MatrixBindings</span></span><br><span data-ttu-id="56df0-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-253">
        - MatrixCoercion</span></span><br><span data-ttu-id="56df0-254">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-254">
        - Selection</span></span><br><span data-ttu-id="56df0-255">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-255">
        - Settings</span></span><br><span data-ttu-id="56df0-256">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-256">
        - TableBindings</span></span><br><span data-ttu-id="56df0-257">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-257">
        - TableCoercion</span></span><br><span data-ttu-id="56df0-258">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-258">
        - TextBindings</span></span><br><span data-ttu-id="56df0-259">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-259">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-260">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="56df0-260">Office apps on Mac</span></span><br><span data-ttu-id="56df0-261">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="56df0-261">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="56df0-262">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-262">- TaskPane</span></span><br><span data-ttu-id="56df0-263">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-263">
        - Content</span></span><br><span data-ttu-id="56df0-264">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="56df0-264">
        - Custom Functions</span></span><br><span data-ttu-id="56df0-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="56df0-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="56df0-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="56df0-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="56df0-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="56df0-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="56df0-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="56df0-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="56df0-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="56df0-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="56df0-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="56df0-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="56df0-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="56df0-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="56df0-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="56df0-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="56df0-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="56df0-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="56df0-278">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-278">- BindingEvents</span></span><br><span data-ttu-id="56df0-279">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-279">
        - CompressedFile</span></span><br><span data-ttu-id="56df0-280">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-280">
        - DocumentEvents</span></span><br><span data-ttu-id="56df0-281">
        - File</span><span class="sxs-lookup"><span data-stu-id="56df0-281">
        - File</span></span><br><span data-ttu-id="56df0-282">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-282">
        - MatrixBindings</span></span><br><span data-ttu-id="56df0-283">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-283">
        - MatrixCoercion</span></span><br><span data-ttu-id="56df0-284">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-284">
        - PdfFile</span></span><br><span data-ttu-id="56df0-285">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-285">
        - Selection</span></span><br><span data-ttu-id="56df0-286">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-286">
        - Settings</span></span><br><span data-ttu-id="56df0-287">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-287">
        - TableBindings</span></span><br><span data-ttu-id="56df0-288">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-288">
        - TableCoercion</span></span><br><span data-ttu-id="56df0-289">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-289">
        - TextBindings</span></span><br><span data-ttu-id="56df0-290">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-290">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-291">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="56df0-291">Office 2019 for Mac</span></span><br><span data-ttu-id="56df0-292">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-292">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="56df0-293">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-293">- TaskPane</span></span><br><span data-ttu-id="56df0-294">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-294">
        - Content</span></span><br><span data-ttu-id="56df0-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="56df0-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="56df0-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="56df0-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="56df0-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="56df0-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="56df0-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="56df0-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="56df0-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="56df0-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="56df0-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="56df0-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="56df0-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="56df0-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="56df0-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="56df0-306">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-306">- BindingEvents</span></span><br><span data-ttu-id="56df0-307">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-307">
        - CompressedFile</span></span><br><span data-ttu-id="56df0-308">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-308">
        - DocumentEvents</span></span><br><span data-ttu-id="56df0-309">
        - File</span><span class="sxs-lookup"><span data-stu-id="56df0-309">
        - File</span></span><br><span data-ttu-id="56df0-310">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-310">
        - MatrixBindings</span></span><br><span data-ttu-id="56df0-311">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-311">
        - MatrixCoercion</span></span><br><span data-ttu-id="56df0-312">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-312">
        - PdfFile</span></span><br><span data-ttu-id="56df0-313">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-313">
        - Selection</span></span><br><span data-ttu-id="56df0-314">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-314">
        - Settings</span></span><br><span data-ttu-id="56df0-315">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-315">
        - TableBindings</span></span><br><span data-ttu-id="56df0-316">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-316">
        - TableCoercion</span></span><br><span data-ttu-id="56df0-317">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-317">
        - TextBindings</span></span><br><span data-ttu-id="56df0-318">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-318">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-319">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="56df0-319">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="56df0-320">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-320">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="56df0-321">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-321">- TaskPane</span></span><br><span data-ttu-id="56df0-322">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-322">
        - Content</span></span></td>
    <td><span data-ttu-id="56df0-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="56df0-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="56df0-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="56df0-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="56df0-326">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-326">- BindingEvents</span></span><br><span data-ttu-id="56df0-327">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-327">
        - CompressedFile</span></span><br><span data-ttu-id="56df0-328">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-328">
        - DocumentEvents</span></span><br><span data-ttu-id="56df0-329">
        - File</span><span class="sxs-lookup"><span data-stu-id="56df0-329">
        - File</span></span><br><span data-ttu-id="56df0-330">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-330">
        - MatrixBindings</span></span><br><span data-ttu-id="56df0-331">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-331">
        - MatrixCoercion</span></span><br><span data-ttu-id="56df0-332">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-332">
        - PdfFile</span></span><br><span data-ttu-id="56df0-333">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-333">
        - Selection</span></span><br><span data-ttu-id="56df0-334">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-334">
        - Settings</span></span><br><span data-ttu-id="56df0-335">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-335">
        - TableBindings</span></span><br><span data-ttu-id="56df0-336">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-336">
        - TableCoercion</span></span><br><span data-ttu-id="56df0-337">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-337">
        - TextBindings</span></span><br><span data-ttu-id="56df0-338">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-338">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="56df0-339">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="56df0-339">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="56df0-340">自定义函数</span><span class="sxs-lookup"><span data-stu-id="56df0-340">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="56df0-341">平台</span><span class="sxs-lookup"><span data-stu-id="56df0-341">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="56df0-342">扩展点</span><span class="sxs-lookup"><span data-stu-id="56df0-342">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="56df0-343">API 要求集</span><span class="sxs-lookup"><span data-stu-id="56df0-343">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="56df0-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="56df0-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-345">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="56df0-345">Office on the web</span></span></td>
    <td><span data-ttu-id="56df0-346">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="56df0-346">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="56df0-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-348">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="56df0-348">Office on Windows</span></span><br><span data-ttu-id="56df0-349">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="56df0-349">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="56df0-350">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="56df0-350">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="56df0-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-352">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="56df0-352">Office for Mac</span></span><br><span data-ttu-id="56df0-353">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="56df0-353">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="56df0-354">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="56df0-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="56df0-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="56df0-356">Outlook</span><span class="sxs-lookup"><span data-stu-id="56df0-356">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="56df0-357">平台</span><span class="sxs-lookup"><span data-stu-id="56df0-357">Platform</span></span></th>
    <th><span data-ttu-id="56df0-358">扩展点</span><span class="sxs-lookup"><span data-stu-id="56df0-358">Extension points</span></span></th>
    <th><span data-ttu-id="56df0-359">API 要求集</span><span class="sxs-lookup"><span data-stu-id="56df0-359">API requirement sets</span></span></th>
    <th><span data-ttu-id="56df0-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="56df0-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-361">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="56df0-361">Office on the web</span></span><br><span data-ttu-id="56df0-362">（新式）</span><span class="sxs-lookup"><span data-stu-id="56df0-362">Modern</span></span></td>
    <td> <span data-ttu-id="56df0-363">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="56df0-363">- Mail Read</span></span><br><span data-ttu-id="56df0-364">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="56df0-364">
      - Mail Compose</span></span><br><span data-ttu-id="56df0-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="56df0-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="56df0-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="56df0-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="56df0-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="56df0-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="56df0-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="56df0-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="56df0-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="56df0-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="56df0-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="56df0-373">不可用</span><span class="sxs-lookup"><span data-stu-id="56df0-373">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-374">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="56df0-374">Office on the web</span></span><br><span data-ttu-id="56df0-375">（经典）</span><span class="sxs-lookup"><span data-stu-id="56df0-375">Classic.</span></span></td>
    <td> <span data-ttu-id="56df0-376">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="56df0-376">- Mail Read</span></span><br><span data-ttu-id="56df0-377">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="56df0-377">
      - Mail Compose</span></span><br><span data-ttu-id="56df0-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="56df0-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="56df0-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="56df0-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="56df0-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="56df0-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="56df0-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="56df0-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="56df0-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="56df0-385">不可用</span><span class="sxs-lookup"><span data-stu-id="56df0-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-386">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="56df0-386">Office on Windows</span></span><br><span data-ttu-id="56df0-387">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="56df0-387">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="56df0-388">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="56df0-388">- Mail Read</span></span><br><span data-ttu-id="56df0-389">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="56df0-389">
      - Mail Compose</span></span><br><span data-ttu-id="56df0-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="56df0-391">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="56df0-391">
      - Modules</span></span></td>
    <td> <span data-ttu-id="56df0-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="56df0-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="56df0-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="56df0-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="56df0-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="56df0-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="56df0-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="56df0-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="56df0-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="56df0-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="56df0-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="56df0-399">不可用</span><span class="sxs-lookup"><span data-stu-id="56df0-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-400">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="56df0-400">Office 2019 on Windows</span></span><br><span data-ttu-id="56df0-401">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-401">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-402">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="56df0-402">- Mail Read</span></span><br><span data-ttu-id="56df0-403">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="56df0-403">
      - Mail Compose</span></span><br><span data-ttu-id="56df0-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="56df0-405">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="56df0-405">
      - Modules</span></span></td>
    <td> <span data-ttu-id="56df0-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="56df0-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="56df0-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="56df0-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="56df0-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="56df0-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="56df0-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="56df0-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="56df0-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="56df0-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="56df0-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="56df0-413">不可用</span><span class="sxs-lookup"><span data-stu-id="56df0-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-414">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="56df0-414">Office 2016 on Windows</span></span><br><span data-ttu-id="56df0-415">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-416">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="56df0-416">- Mail Read</span></span><br><span data-ttu-id="56df0-417">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="56df0-417">
      - Mail Compose</span></span><br><span data-ttu-id="56df0-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="56df0-419">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="56df0-419">
      - Modules</span></span></td>
    <td> <span data-ttu-id="56df0-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="56df0-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="56df0-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="56df0-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="56df0-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="56df0-424">不可用</span><span class="sxs-lookup"><span data-stu-id="56df0-424">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-425">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="56df0-425">Office 2013 on Windows</span></span><br><span data-ttu-id="56df0-426">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-426">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-427">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="56df0-427">- Mail Read</span></span><br><span data-ttu-id="56df0-428">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="56df0-428">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="56df0-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="56df0-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="56df0-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="56df0-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="56df0-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="56df0-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="56df0-433">不可用</span><span class="sxs-lookup"><span data-stu-id="56df0-433">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-434">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="56df0-434">Office apps on iOS</span></span><br><span data-ttu-id="56df0-435">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="56df0-435">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="56df0-436">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="56df0-436">- Mail Read</span></span><br><span data-ttu-id="56df0-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="56df0-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="56df0-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="56df0-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="56df0-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="56df0-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="56df0-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="56df0-443">不可用</span><span class="sxs-lookup"><span data-stu-id="56df0-443">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-444">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="56df0-444">Office apps on Mac</span></span><br><span data-ttu-id="56df0-445">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="56df0-445">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="56df0-446">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="56df0-446">- Mail Read</span></span><br><span data-ttu-id="56df0-447">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="56df0-447">
      - Mail Compose</span></span><br><span data-ttu-id="56df0-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="56df0-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="56df0-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="56df0-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="56df0-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="56df0-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="56df0-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="56df0-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="56df0-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="56df0-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="56df0-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="56df0-456">不可用</span><span class="sxs-lookup"><span data-stu-id="56df0-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-457">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="56df0-457">Office 2019 for Mac</span></span><br><span data-ttu-id="56df0-458">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-458">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-459">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="56df0-459">- Mail Read</span></span><br><span data-ttu-id="56df0-460">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="56df0-460">
      - Mail Compose</span></span><br><span data-ttu-id="56df0-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="56df0-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="56df0-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="56df0-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="56df0-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="56df0-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="56df0-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="56df0-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="56df0-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="56df0-468">不可用</span><span class="sxs-lookup"><span data-stu-id="56df0-468">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-469">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="56df0-469">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="56df0-470">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-470">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-471">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="56df0-471">- Mail Read</span></span><br><span data-ttu-id="56df0-472">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="56df0-472">
      - Mail Compose</span></span><br><span data-ttu-id="56df0-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="56df0-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="56df0-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="56df0-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="56df0-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="56df0-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="56df0-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="56df0-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="56df0-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="56df0-480">不可用</span><span class="sxs-lookup"><span data-stu-id="56df0-480">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-481">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="56df0-481">Office apps on Android</span></span><br><span data-ttu-id="56df0-482">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="56df0-482">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="56df0-483">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="56df0-483">- Mail Read</span></span><br><span data-ttu-id="56df0-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="56df0-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="56df0-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="56df0-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="56df0-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="56df0-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="56df0-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="56df0-490">不可用</span><span class="sxs-lookup"><span data-stu-id="56df0-490">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="56df0-491">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="56df0-491">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="56df0-492">Word</span><span class="sxs-lookup"><span data-stu-id="56df0-492">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="56df0-493">平台</span><span class="sxs-lookup"><span data-stu-id="56df0-493">Platform</span></span></th>
    <th><span data-ttu-id="56df0-494">扩展点</span><span class="sxs-lookup"><span data-stu-id="56df0-494">Extension points</span></span></th>
    <th><span data-ttu-id="56df0-495">API 要求集</span><span class="sxs-lookup"><span data-stu-id="56df0-495">API requirement sets</span></span></th>
    <th><span data-ttu-id="56df0-496"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="56df0-496"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-497">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="56df0-497">Office on the web</span></span></td>
    <td> <span data-ttu-id="56df0-498">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-498">- TaskPane</span></span><br><span data-ttu-id="56df0-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-500">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-500">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="56df0-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="56df0-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="56df0-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="56df0-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="56df0-506">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-506">- BindingEvents</span></span><br><span data-ttu-id="56df0-507">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="56df0-507">
         - CustomXmlParts</span></span><br><span data-ttu-id="56df0-508">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-508">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-509">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-509">
         - File</span></span><br><span data-ttu-id="56df0-510">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-510">
         - HtmlCoercion</span></span><br><span data-ttu-id="56df0-511">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-511">
         - MatrixBindings</span></span><br><span data-ttu-id="56df0-512">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-512">
         - MatrixCoercion</span></span><br><span data-ttu-id="56df0-513">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-513">
         - OoxmlCoercion</span></span><br><span data-ttu-id="56df0-514">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-514">
         - PdfFile</span></span><br><span data-ttu-id="56df0-515">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-515">
         - Selection</span></span><br><span data-ttu-id="56df0-516">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-516">
         - Settings</span></span><br><span data-ttu-id="56df0-517">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-517">
         - TableBindings</span></span><br><span data-ttu-id="56df0-518">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-518">
         - TableCoercion</span></span><br><span data-ttu-id="56df0-519">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-519">
         - TextBindings</span></span><br><span data-ttu-id="56df0-520">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-520">
         - TextCoercion</span></span><br><span data-ttu-id="56df0-521">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="56df0-521">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-522">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="56df0-522">Office on Windows</span></span><br><span data-ttu-id="56df0-523">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="56df0-523">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="56df0-524">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-524">- TaskPane</span></span><br><span data-ttu-id="56df0-525">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-525">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-526">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-526">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="56df0-527">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-527">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="56df0-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="56df0-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-530">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-530">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="56df0-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="56df0-532">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-532">- BindingEvents</span></span><br><span data-ttu-id="56df0-533">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-533">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-534">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="56df0-534">
         - CustomXmlParts</span></span><br><span data-ttu-id="56df0-535">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-535">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-536">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-536">
         - File</span></span><br><span data-ttu-id="56df0-537">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-537">
         - HtmlCoercion</span></span><br><span data-ttu-id="56df0-538">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-538">
         - MatrixBindings</span></span><br><span data-ttu-id="56df0-539">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-539">
         - MatrixCoercion</span></span><br><span data-ttu-id="56df0-540">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-540">
         - OoxmlCoercion</span></span><br><span data-ttu-id="56df0-541">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-541">
         - PdfFile</span></span><br><span data-ttu-id="56df0-542">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-542">
         - Selection</span></span><br><span data-ttu-id="56df0-543">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-543">
         - Settings</span></span><br><span data-ttu-id="56df0-544">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-544">
         - TableBindings</span></span><br><span data-ttu-id="56df0-545">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-545">
         - TableCoercion</span></span><br><span data-ttu-id="56df0-546">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-546">
         - TextBindings</span></span><br><span data-ttu-id="56df0-547">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-547">
         - TextCoercion</span></span><br><span data-ttu-id="56df0-548">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="56df0-548">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-549">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="56df0-549">Office 2019 on Windows</span></span><br><span data-ttu-id="56df0-550">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-550">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-551">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-551">- TaskPane</span></span><br><span data-ttu-id="56df0-552">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-552">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-553">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-553">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="56df0-554">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-554">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="56df0-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="56df0-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-557">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-557">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="56df0-558">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-558">- BindingEvents</span></span><br><span data-ttu-id="56df0-559">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-559">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-560">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="56df0-560">
         - CustomXmlParts</span></span><br><span data-ttu-id="56df0-561">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-561">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-562">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-562">
         - File</span></span><br><span data-ttu-id="56df0-563">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-563">
         - HtmlCoercion</span></span><br><span data-ttu-id="56df0-564">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-564">
         - MatrixBindings</span></span><br><span data-ttu-id="56df0-565">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-565">
         - MatrixCoercion</span></span><br><span data-ttu-id="56df0-566">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-566">
         - OoxmlCoercion</span></span><br><span data-ttu-id="56df0-567">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-567">
         - PdfFile</span></span><br><span data-ttu-id="56df0-568">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-568">
         - Selection</span></span><br><span data-ttu-id="56df0-569">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-569">
         - Settings</span></span><br><span data-ttu-id="56df0-570">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-570">
         - TableBindings</span></span><br><span data-ttu-id="56df0-571">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-571">
         - TableCoercion</span></span><br><span data-ttu-id="56df0-572">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-572">
         - TextBindings</span></span><br><span data-ttu-id="56df0-573">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-573">
         - TextCoercion</span></span><br><span data-ttu-id="56df0-574">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="56df0-574">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-575">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="56df0-575">Office 2016 on Windows</span></span><br><span data-ttu-id="56df0-576">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-576">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-577">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-577">- TaskPane</span></span></td>
    <td> <span data-ttu-id="56df0-578">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-578">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="56df0-579">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="56df0-579">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="56df0-580">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-580">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="56df0-581">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-581">- BindingEvents</span></span><br><span data-ttu-id="56df0-582">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-582">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-583">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="56df0-583">
         - CustomXmlParts</span></span><br><span data-ttu-id="56df0-584">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-584">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-585">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-585">
         - File</span></span><br><span data-ttu-id="56df0-586">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-586">
         - HtmlCoercion</span></span><br><span data-ttu-id="56df0-587">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-587">
         - MatrixBindings</span></span><br><span data-ttu-id="56df0-588">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-588">
         - MatrixCoercion</span></span><br><span data-ttu-id="56df0-589">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-589">
         - OoxmlCoercion</span></span><br><span data-ttu-id="56df0-590">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-590">
         - PdfFile</span></span><br><span data-ttu-id="56df0-591">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-591">
         - Selection</span></span><br><span data-ttu-id="56df0-592">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-592">
         - Settings</span></span><br><span data-ttu-id="56df0-593">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-593">
         - TableBindings</span></span><br><span data-ttu-id="56df0-594">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-594">
         - TableCoercion</span></span><br><span data-ttu-id="56df0-595">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-595">
         - TextBindings</span></span><br><span data-ttu-id="56df0-596">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-596">
         - TextCoercion</span></span><br><span data-ttu-id="56df0-597">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="56df0-597">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-598">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="56df0-598">Office 2013 on Windows</span></span><br><span data-ttu-id="56df0-599">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-599">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-600">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-600">- TaskPane</span></span></td>
    <td> <span data-ttu-id="56df0-601">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="56df0-601">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="56df0-602">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-602">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="56df0-603">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-603">- BindingEvents</span></span><br><span data-ttu-id="56df0-604">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-604">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-605">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="56df0-605">
         - CustomXmlParts</span></span><br><span data-ttu-id="56df0-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-606">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-607">
         - File</span></span><br><span data-ttu-id="56df0-608">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-608">
         - HtmlCoercion</span></span><br><span data-ttu-id="56df0-609">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-609">
         - MatrixBindings</span></span><br><span data-ttu-id="56df0-610">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-610">
         - MatrixCoercion</span></span><br><span data-ttu-id="56df0-611">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-611">
         - OoxmlCoercion</span></span><br><span data-ttu-id="56df0-612">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-612">
         - PdfFile</span></span><br><span data-ttu-id="56df0-613">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-613">
         - Selection</span></span><br><span data-ttu-id="56df0-614">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-614">
         - Settings</span></span><br><span data-ttu-id="56df0-615">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-615">
         - TableBindings</span></span><br><span data-ttu-id="56df0-616">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-616">
         - TableCoercion</span></span><br><span data-ttu-id="56df0-617">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-617">
         - TextBindings</span></span><br><span data-ttu-id="56df0-618">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-618">
         - TextCoercion</span></span><br><span data-ttu-id="56df0-619">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="56df0-619">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-620">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="56df0-620">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="56df0-621">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="56df0-621">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="56df0-622">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-622">- TaskPane</span></span></td>
    <td> <span data-ttu-id="56df0-623">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-623">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="56df0-624">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-624">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="56df0-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="56df0-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="56df0-628">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-628">- BindingEvents</span></span><br><span data-ttu-id="56df0-629">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-629">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-630">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="56df0-630">
         - CustomXmlParts</span></span><br><span data-ttu-id="56df0-631">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-631">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-632">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-632">
         - File</span></span><br><span data-ttu-id="56df0-633">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-633">
         - HtmlCoercion</span></span><br><span data-ttu-id="56df0-634">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-634">
         - MatrixBindings</span></span><br><span data-ttu-id="56df0-635">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-635">
         - MatrixCoercion</span></span><br><span data-ttu-id="56df0-636">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-636">
         - OoxmlCoercion</span></span><br><span data-ttu-id="56df0-637">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-637">
         - PdfFile</span></span><br><span data-ttu-id="56df0-638">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-638">
         - Selection</span></span><br><span data-ttu-id="56df0-639">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-639">
         - Settings</span></span><br><span data-ttu-id="56df0-640">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-640">
         - TableBindings</span></span><br><span data-ttu-id="56df0-641">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-641">
         - TableCoercion</span></span><br><span data-ttu-id="56df0-642">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-642">
         - TextBindings</span></span><br><span data-ttu-id="56df0-643">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-643">
         - TextCoercion</span></span><br><span data-ttu-id="56df0-644">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="56df0-644">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-645">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="56df0-645">Office apps on Mac</span></span><br><span data-ttu-id="56df0-646">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="56df0-646">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="56df0-647">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-647">- TaskPane</span></span><br><span data-ttu-id="56df0-648">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-648">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-649">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-649">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="56df0-650">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-650">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="56df0-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="56df0-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="56df0-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="56df0-655">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-655">- BindingEvents</span></span><br><span data-ttu-id="56df0-656">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-656">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-657">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="56df0-657">
         - CustomXmlParts</span></span><br><span data-ttu-id="56df0-658">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-658">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-659">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-659">
         - File</span></span><br><span data-ttu-id="56df0-660">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-660">
         - HtmlCoercion</span></span><br><span data-ttu-id="56df0-661">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-661">
         - MatrixBindings</span></span><br><span data-ttu-id="56df0-662">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-662">
         - MatrixCoercion</span></span><br><span data-ttu-id="56df0-663">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-663">
         - OoxmlCoercion</span></span><br><span data-ttu-id="56df0-664">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-664">
         - PdfFile</span></span><br><span data-ttu-id="56df0-665">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-665">
         - Selection</span></span><br><span data-ttu-id="56df0-666">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-666">
         - Settings</span></span><br><span data-ttu-id="56df0-667">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-667">
         - TableBindings</span></span><br><span data-ttu-id="56df0-668">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-668">
         - TableCoercion</span></span><br><span data-ttu-id="56df0-669">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-669">
         - TextBindings</span></span><br><span data-ttu-id="56df0-670">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-670">
         - TextCoercion</span></span><br><span data-ttu-id="56df0-671">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="56df0-671">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-672">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="56df0-672">Office 2019 for Mac</span></span><br><span data-ttu-id="56df0-673">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-673">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-674">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-674">- TaskPane</span></span><br><span data-ttu-id="56df0-675">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-675">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-676">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-676">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="56df0-677">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-677">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="56df0-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="56df0-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="56df0-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="56df0-681">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-681">- BindingEvents</span></span><br><span data-ttu-id="56df0-682">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-682">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-683">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="56df0-683">
         - CustomXmlParts</span></span><br><span data-ttu-id="56df0-684">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-684">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-685">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-685">
         - File</span></span><br><span data-ttu-id="56df0-686">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-686">
         - HtmlCoercion</span></span><br><span data-ttu-id="56df0-687">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-687">
         - MatrixBindings</span></span><br><span data-ttu-id="56df0-688">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-688">
         - MatrixCoercion</span></span><br><span data-ttu-id="56df0-689">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-689">
         - OoxmlCoercion</span></span><br><span data-ttu-id="56df0-690">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-690">
         - PdfFile</span></span><br><span data-ttu-id="56df0-691">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-691">
         - Selection</span></span><br><span data-ttu-id="56df0-692">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-692">
         - Settings</span></span><br><span data-ttu-id="56df0-693">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-693">
         - TableBindings</span></span><br><span data-ttu-id="56df0-694">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-694">
         - TableCoercion</span></span><br><span data-ttu-id="56df0-695">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-695">
         - TextBindings</span></span><br><span data-ttu-id="56df0-696">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-696">
         - TextCoercion</span></span><br><span data-ttu-id="56df0-697">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="56df0-697">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-698">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="56df0-698">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="56df0-699">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-699">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-700">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-700">- TaskPane</span></span></td>
    <td> <span data-ttu-id="56df0-701">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-701">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="56df0-702">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="56df0-702">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="56df0-703">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-703">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="56df0-704">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-704">- BindingEvents</span></span><br><span data-ttu-id="56df0-705">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-705">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-706">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="56df0-706">
         - CustomXmlParts</span></span><br><span data-ttu-id="56df0-707">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-707">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-708">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-708">
         - File</span></span><br><span data-ttu-id="56df0-709">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-709">
         - HtmlCoercion</span></span><br><span data-ttu-id="56df0-710">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-710">
         - MatrixBindings</span></span><br><span data-ttu-id="56df0-711">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-711">
         - MatrixCoercion</span></span><br><span data-ttu-id="56df0-712">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-712">
         - OoxmlCoercion</span></span><br><span data-ttu-id="56df0-713">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-713">
         - PdfFile</span></span><br><span data-ttu-id="56df0-714">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-714">
         - Selection</span></span><br><span data-ttu-id="56df0-715">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-715">
         - Settings</span></span><br><span data-ttu-id="56df0-716">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-716">
         - TableBindings</span></span><br><span data-ttu-id="56df0-717">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-717">
         - TableCoercion</span></span><br><span data-ttu-id="56df0-718">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="56df0-718">
         - TextBindings</span></span><br><span data-ttu-id="56df0-719">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-719">
         - TextCoercion</span></span><br><span data-ttu-id="56df0-720">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="56df0-720">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="56df0-721">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="56df0-721">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="56df0-722">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="56df0-722">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="56df0-723">平台</span><span class="sxs-lookup"><span data-stu-id="56df0-723">Platform</span></span></th>
    <th><span data-ttu-id="56df0-724">扩展点</span><span class="sxs-lookup"><span data-stu-id="56df0-724">Extension points</span></span></th>
    <th><span data-ttu-id="56df0-725">API 要求集</span><span class="sxs-lookup"><span data-stu-id="56df0-725">API requirement sets</span></span></th>
    <th><span data-ttu-id="56df0-726"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="56df0-726"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-727">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="56df0-727">Office on the web</span></span></td>
    <td> <span data-ttu-id="56df0-728">- 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-728">- Content</span></span><br><span data-ttu-id="56df0-729">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-729">
         - TaskPane</span></span><br><span data-ttu-id="56df0-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-731">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-731">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="56df0-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="56df0-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="56df0-735">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="56df0-735">- ActiveView</span></span><br><span data-ttu-id="56df0-736">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-736">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-737">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-737">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-738">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-738">
         - File</span></span><br><span data-ttu-id="56df0-739">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-739">
         - PdfFile</span></span><br><span data-ttu-id="56df0-740">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-740">
         - Selection</span></span><br><span data-ttu-id="56df0-741">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-741">
         - Settings</span></span><br><span data-ttu-id="56df0-742">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-742">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-743">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="56df0-743">Office on Windows</span></span><br><span data-ttu-id="56df0-744">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="56df0-744">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="56df0-745">- 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-745">- Content</span></span><br><span data-ttu-id="56df0-746">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-746">
         - TaskPane</span></span><br><span data-ttu-id="56df0-747">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-747">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-748">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-748">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="56df0-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="56df0-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="56df0-752">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="56df0-752">- ActiveView</span></span><br><span data-ttu-id="56df0-753">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-753">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-754">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-754">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-755">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-755">
         - File</span></span><br><span data-ttu-id="56df0-756">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-756">
         - PdfFile</span></span><br><span data-ttu-id="56df0-757">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-757">
         - Selection</span></span><br><span data-ttu-id="56df0-758">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-758">
         - Settings</span></span><br><span data-ttu-id="56df0-759">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-759">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-760">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="56df0-760">Office 2019 on Windows</span></span><br><span data-ttu-id="56df0-761">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-761">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-762">- 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-762">- Content</span></span><br><span data-ttu-id="56df0-763">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-763">
         - TaskPane</span></span><br><span data-ttu-id="56df0-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-765">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-765">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="56df0-767">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="56df0-767">- ActiveView</span></span><br><span data-ttu-id="56df0-768">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-768">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-769">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-769">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-770">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-770">
         - File</span></span><br><span data-ttu-id="56df0-771">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-771">
         - PdfFile</span></span><br><span data-ttu-id="56df0-772">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-772">
         - Selection</span></span><br><span data-ttu-id="56df0-773">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-773">
         - Settings</span></span><br><span data-ttu-id="56df0-774">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-774">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-775">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="56df0-775">Office 2016 on Windows</span></span><br><span data-ttu-id="56df0-776">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-776">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-777">- 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-777">- Content</span></span><br><span data-ttu-id="56df0-778">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-778">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="56df0-779">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="56df0-779">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="56df0-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="56df0-781">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="56df0-781">- ActiveView</span></span><br><span data-ttu-id="56df0-782">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-782">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-783">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-783">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-784">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-784">
         - File</span></span><br><span data-ttu-id="56df0-785">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-785">
         - PdfFile</span></span><br><span data-ttu-id="56df0-786">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-786">
         - Selection</span></span><br><span data-ttu-id="56df0-787">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-787">
         - Settings</span></span><br><span data-ttu-id="56df0-788">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-788">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-789">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="56df0-789">Office 2013 on Windows</span></span><br><span data-ttu-id="56df0-790">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-790">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-791">- 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-791">- Content</span></span><br><span data-ttu-id="56df0-792">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-792">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="56df0-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="56df0-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="56df0-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="56df0-795">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="56df0-795">- ActiveView</span></span><br><span data-ttu-id="56df0-796">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-796">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-797">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-797">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-798">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-798">
         - File</span></span><br><span data-ttu-id="56df0-799">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-799">
         - PdfFile</span></span><br><span data-ttu-id="56df0-800">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-800">
         - Selection</span></span><br><span data-ttu-id="56df0-801">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-801">
         - Settings</span></span><br><span data-ttu-id="56df0-802">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-802">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-803">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="56df0-803">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="56df0-804">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="56df0-804">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="56df0-805">- 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-805">- Content</span></span><br><span data-ttu-id="56df0-806">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-806">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="56df0-807">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-807">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="56df0-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="56df0-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="56df0-810">- ActiveView</span></span><br><span data-ttu-id="56df0-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-811">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-812">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-813">
         - File</span></span><br><span data-ttu-id="56df0-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-814">
         - PdfFile</span></span><br><span data-ttu-id="56df0-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-815">
         - Selection</span></span><br><span data-ttu-id="56df0-816">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-816">
         - Settings</span></span><br><span data-ttu-id="56df0-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-818">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="56df0-818">Office apps on Mac</span></span><br><span data-ttu-id="56df0-819">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="56df0-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="56df0-820">- 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-820">- Content</span></span><br><span data-ttu-id="56df0-821">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-821">
         - TaskPane</span></span><br><span data-ttu-id="56df0-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-823">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-823">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="56df0-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="56df0-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="56df0-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="56df0-827">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="56df0-827">- ActiveView</span></span><br><span data-ttu-id="56df0-828">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-828">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-829">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-829">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-830">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-830">
         - File</span></span><br><span data-ttu-id="56df0-831">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-831">
         - PdfFile</span></span><br><span data-ttu-id="56df0-832">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-832">
         - Selection</span></span><br><span data-ttu-id="56df0-833">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-833">
         - Settings</span></span><br><span data-ttu-id="56df0-834">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-834">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-835">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="56df0-835">Office 2019 for Mac</span></span><br><span data-ttu-id="56df0-836">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-836">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-837">- 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-837">- Content</span></span><br><span data-ttu-id="56df0-838">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-838">
         - TaskPane</span></span><br><span data-ttu-id="56df0-839">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-839">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-840">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-840">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="56df0-842">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="56df0-842">- ActiveView</span></span><br><span data-ttu-id="56df0-843">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-843">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-844">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-844">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-845">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-845">
         - File</span></span><br><span data-ttu-id="56df0-846">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-846">
         - PdfFile</span></span><br><span data-ttu-id="56df0-847">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-847">
         - Selection</span></span><br><span data-ttu-id="56df0-848">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-848">
         - Settings</span></span><br><span data-ttu-id="56df0-849">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-849">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-850">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="56df0-850">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="56df0-851">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-851">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-852">- 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-852">- Content</span></span><br><span data-ttu-id="56df0-853">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-853">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="56df0-854">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="56df0-854">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="56df0-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="56df0-856">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="56df0-856">- ActiveView</span></span><br><span data-ttu-id="56df0-857">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="56df0-857">
         - CompressedFile</span></span><br><span data-ttu-id="56df0-858">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-858">
         - DocumentEvents</span></span><br><span data-ttu-id="56df0-859">
         - File</span><span class="sxs-lookup"><span data-stu-id="56df0-859">
         - File</span></span><br><span data-ttu-id="56df0-860">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="56df0-860">
         - PdfFile</span></span><br><span data-ttu-id="56df0-861">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-861">
         - Selection</span></span><br><span data-ttu-id="56df0-862">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-862">
         - Settings</span></span><br><span data-ttu-id="56df0-863">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-863">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="56df0-864">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="56df0-864">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="56df0-865">OneNote</span><span class="sxs-lookup"><span data-stu-id="56df0-865">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="56df0-866">平台</span><span class="sxs-lookup"><span data-stu-id="56df0-866">Platform</span></span></th>
    <th><span data-ttu-id="56df0-867">扩展点</span><span class="sxs-lookup"><span data-stu-id="56df0-867">Extension points</span></span></th>
    <th><span data-ttu-id="56df0-868">API 要求集</span><span class="sxs-lookup"><span data-stu-id="56df0-868">API requirement sets</span></span></th>
    <th><span data-ttu-id="56df0-869"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="56df0-869"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-870">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="56df0-870">Office on the web</span></span></td>
    <td> <span data-ttu-id="56df0-871">- 内容</span><span class="sxs-lookup"><span data-stu-id="56df0-871">- Content</span></span><br><span data-ttu-id="56df0-872">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-872">
         - TaskPane</span></span><br><span data-ttu-id="56df0-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="56df0-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="56df0-874">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-874">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="56df0-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="56df0-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="56df0-877">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="56df0-877">- DocumentEvents</span></span><br><span data-ttu-id="56df0-878">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-878">
         - HtmlCoercion</span></span><br><span data-ttu-id="56df0-879">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="56df0-879">
         - Settings</span></span><br><span data-ttu-id="56df0-880">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-880">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="56df0-881">项目</span><span class="sxs-lookup"><span data-stu-id="56df0-881">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="56df0-882">平台</span><span class="sxs-lookup"><span data-stu-id="56df0-882">Platform</span></span></th>
    <th><span data-ttu-id="56df0-883">扩展点</span><span class="sxs-lookup"><span data-stu-id="56df0-883">Extension points</span></span></th>
    <th><span data-ttu-id="56df0-884">API 要求集</span><span class="sxs-lookup"><span data-stu-id="56df0-884">API requirement sets</span></span></th>
    <th><span data-ttu-id="56df0-885"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="56df0-885"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-886">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="56df0-886">Office 2019 on Windows</span></span><br><span data-ttu-id="56df0-887">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-887">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-888">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-888">- TaskPane</span></span></td>
    <td> <span data-ttu-id="56df0-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="56df0-890">- Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-890">- Selection</span></span><br><span data-ttu-id="56df0-891">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-891">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-892">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="56df0-892">Office 2016 on Windows</span></span><br><span data-ttu-id="56df0-893">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-893">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-894">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-894">- TaskPane</span></span></td>
    <td> <span data-ttu-id="56df0-895">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-895">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="56df0-896">- Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-896">- Selection</span></span><br><span data-ttu-id="56df0-897">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-897">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="56df0-898">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="56df0-898">Office 2013 on Windows</span></span><br><span data-ttu-id="56df0-899">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="56df0-899">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="56df0-900">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="56df0-900">- TaskPane</span></span></td>
    <td> <span data-ttu-id="56df0-901">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="56df0-901">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="56df0-902">- Selection</span><span class="sxs-lookup"><span data-stu-id="56df0-902">- Selection</span></span><br><span data-ttu-id="56df0-903">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="56df0-903">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="56df0-904">另请参阅</span><span class="sxs-lookup"><span data-stu-id="56df0-904">See also</span></span>

- [<span data-ttu-id="56df0-905">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="56df0-905">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="56df0-906">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="56df0-906">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="56df0-907">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="56df0-907">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="56df0-908">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="56df0-908">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="56df0-909">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="56df0-909">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="56df0-910">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="56df0-910">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="56df0-911">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="56df0-911">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="56df0-912">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="56df0-912">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="56df0-913">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="56df0-913">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="56df0-914">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="56df0-914">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="56df0-915">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="56df0-915">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)

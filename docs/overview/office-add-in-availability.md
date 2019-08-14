---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 08/13/2019
localization_priority: Priority
ms.openlocfilehash: 1e368fe21a1bcdb2a7f44c88ce8e881605fa96f2
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395650"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="12b5d-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="12b5d-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="12b5d-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="12b5d-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="12b5d-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="12b5d-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="12b5d-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="12b5d-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="12b5d-108">Excel</span><span class="sxs-lookup"><span data-stu-id="12b5d-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="12b5d-109">平台</span><span class="sxs-lookup"><span data-stu-id="12b5d-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="12b5d-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="12b5d-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="12b5d-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="12b5d-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="12b5d-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="12b5d-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="12b5d-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="12b5d-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-114">- TaskPane</span></span><br><span data-ttu-id="12b5d-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-115">
        - Content</span></span><br><span data-ttu-id="12b5d-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="12b5d-116">
        - Custom Functions</span></span><br><span data-ttu-id="12b5d-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="12b5d-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="12b5d-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12b5d-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12b5d-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12b5d-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12b5d-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12b5d-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12b5d-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12b5d-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12b5d-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="12b5d-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="12b5d-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-128">
        - BindingEvents</span></span><br><span data-ttu-id="12b5d-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-129">
        - CompressedFile</span></span><br><span data-ttu-id="12b5d-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-130">
        - DocumentEvents</span></span><br><span data-ttu-id="12b5d-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-131">
        - File</span></span><br><span data-ttu-id="12b5d-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-132">
        - MatrixBindings</span></span><br><span data-ttu-id="12b5d-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-134">
        - Selection</span></span><br><span data-ttu-id="12b5d-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-135">
        - Settings</span></span><br><span data-ttu-id="12b5d-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-136">
        - TableBindings</span></span><br><span data-ttu-id="12b5d-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-137">
        - TableCoercion</span></span><br><span data-ttu-id="12b5d-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-138">
        - TextBindings</span></span><br><span data-ttu-id="12b5d-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-140">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="12b5d-140">Office on Windows</span></span><br><span data-ttu-id="12b5d-141">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="12b5d-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="12b5d-142">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-142">- TaskPane</span></span><br><span data-ttu-id="12b5d-143">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-143">
        - Content</span></span><br><span data-ttu-id="12b5d-144">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="12b5d-144">
        - Custom Functions</span></span><br><span data-ttu-id="12b5d-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="12b5d-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="12b5d-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12b5d-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12b5d-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12b5d-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12b5d-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12b5d-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12b5d-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12b5d-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12b5d-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="12b5d-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="12b5d-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="12b5d-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-158">
        - BindingEvents</span></span><br><span data-ttu-id="12b5d-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-159">
        - CompressedFile</span></span><br><span data-ttu-id="12b5d-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-160">
        - DocumentEvents</span></span><br><span data-ttu-id="12b5d-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-161">
        - File</span></span><br><span data-ttu-id="12b5d-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-162">
        - MatrixBindings</span></span><br><span data-ttu-id="12b5d-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-164">
        - Selection</span></span><br><span data-ttu-id="12b5d-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-165">
        - Settings</span></span><br><span data-ttu-id="12b5d-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-166">
        - TableBindings</span></span><br><span data-ttu-id="12b5d-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-167">
        - TableCoercion</span></span><br><span data-ttu-id="12b5d-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-168">
        - TextBindings</span></span><br><span data-ttu-id="12b5d-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-170">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="12b5d-170">Office 2019 on Windows</span></span><br><span data-ttu-id="12b5d-171">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="12b5d-172">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-172">- TaskPane</span></span><br><span data-ttu-id="12b5d-173">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-173">
        - Content</span></span><br><span data-ttu-id="12b5d-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="12b5d-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12b5d-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12b5d-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12b5d-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12b5d-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12b5d-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12b5d-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12b5d-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12b5d-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="12b5d-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-185">- BindingEvents</span></span><br><span data-ttu-id="12b5d-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-186">
        - CompressedFile</span></span><br><span data-ttu-id="12b5d-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-187">
        - DocumentEvents</span></span><br><span data-ttu-id="12b5d-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-188">
        - File</span></span><br><span data-ttu-id="12b5d-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-189">
        - MatrixBindings</span></span><br><span data-ttu-id="12b5d-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-191">
        - Selection</span></span><br><span data-ttu-id="12b5d-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-192">
        - Settings</span></span><br><span data-ttu-id="12b5d-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-193">
        - TableBindings</span></span><br><span data-ttu-id="12b5d-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-194">
        - TableCoercion</span></span><br><span data-ttu-id="12b5d-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-195">
        - TextBindings</span></span><br><span data-ttu-id="12b5d-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-197">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="12b5d-197">Office 2016 on Windows</span></span><br><span data-ttu-id="12b5d-198">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="12b5d-199">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-199">- TaskPane</span></span><br><span data-ttu-id="12b5d-200">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-200">
        - Content</span></span></td>
    <td><span data-ttu-id="12b5d-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12b5d-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="12b5d-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="12b5d-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="12b5d-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-204">- BindingEvents</span></span><br><span data-ttu-id="12b5d-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-205">
        - CompressedFile</span></span><br><span data-ttu-id="12b5d-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-206">
        - DocumentEvents</span></span><br><span data-ttu-id="12b5d-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-207">
        - File</span></span><br><span data-ttu-id="12b5d-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-208">
        - MatrixBindings</span></span><br><span data-ttu-id="12b5d-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-210">
        - Selection</span></span><br><span data-ttu-id="12b5d-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-211">
        - Settings</span></span><br><span data-ttu-id="12b5d-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-212">
        - TableBindings</span></span><br><span data-ttu-id="12b5d-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-213">
        - TableCoercion</span></span><br><span data-ttu-id="12b5d-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-214">
        - TextBindings</span></span><br><span data-ttu-id="12b5d-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-216">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="12b5d-216">Office 2013 on Windows</span></span><br><span data-ttu-id="12b5d-217">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="12b5d-218">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-218">
        - TaskPane</span></span><br><span data-ttu-id="12b5d-219">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="12b5d-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="12b5d-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="12b5d-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="12b5d-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-222">
        - BindingEvents</span></span><br><span data-ttu-id="12b5d-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-223">
        - CompressedFile</span></span><br><span data-ttu-id="12b5d-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-224">
        - DocumentEvents</span></span><br><span data-ttu-id="12b5d-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-225">
        - File</span></span><br><span data-ttu-id="12b5d-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-226">
        - MatrixBindings</span></span><br><span data-ttu-id="12b5d-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-228">
        - Selection</span></span><br><span data-ttu-id="12b5d-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-229">
        - Settings</span></span><br><span data-ttu-id="12b5d-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-230">
        - TableBindings</span></span><br><span data-ttu-id="12b5d-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-231">
        - TableCoercion</span></span><br><span data-ttu-id="12b5d-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-232">
        - TextBindings</span></span><br><span data-ttu-id="12b5d-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-234">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="12b5d-234">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="12b5d-235">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="12b5d-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="12b5d-236">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-236">- TaskPane</span></span><br><span data-ttu-id="12b5d-237">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-237">
        - Content</span></span><br><span data-ttu-id="12b5d-238">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="12b5d-238">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="12b5d-239">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-239">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12b5d-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12b5d-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12b5d-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12b5d-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12b5d-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12b5d-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12b5d-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12b5d-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="12b5d-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="12b5d-250">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-250">- BindingEvents</span></span><br><span data-ttu-id="12b5d-251">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-251">
        - DocumentEvents</span></span><br><span data-ttu-id="12b5d-252">
        - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-252">
        - File</span></span><br><span data-ttu-id="12b5d-253">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-253">
        - MatrixBindings</span></span><br><span data-ttu-id="12b5d-254">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-254">
        - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-255">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-255">
        - Selection</span></span><br><span data-ttu-id="12b5d-256">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-256">
        - Settings</span></span><br><span data-ttu-id="12b5d-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-257">
        - TableBindings</span></span><br><span data-ttu-id="12b5d-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-258">
        - TableCoercion</span></span><br><span data-ttu-id="12b5d-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-259">
        - TextBindings</span></span><br><span data-ttu-id="12b5d-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-260">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-261">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="12b5d-261">Office apps on Mac</span></span><br><span data-ttu-id="12b5d-262">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="12b5d-262">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="12b5d-263">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-263">- TaskPane</span></span><br><span data-ttu-id="12b5d-264">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-264">
        - Content</span></span><br><span data-ttu-id="12b5d-265">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="12b5d-265">
        - Custom Functions</span></span><br><span data-ttu-id="12b5d-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="12b5d-267">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-267">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12b5d-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12b5d-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12b5d-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12b5d-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12b5d-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12b5d-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12b5d-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12b5d-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="12b5d-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="12b5d-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="12b5d-279">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-279">- BindingEvents</span></span><br><span data-ttu-id="12b5d-280">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-280">
        - CompressedFile</span></span><br><span data-ttu-id="12b5d-281">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-281">
        - DocumentEvents</span></span><br><span data-ttu-id="12b5d-282">
        - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-282">
        - File</span></span><br><span data-ttu-id="12b5d-283">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-283">
        - MatrixBindings</span></span><br><span data-ttu-id="12b5d-284">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-284">
        - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-285">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-285">
        - PdfFile</span></span><br><span data-ttu-id="12b5d-286">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-286">
        - Selection</span></span><br><span data-ttu-id="12b5d-287">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-287">
        - Settings</span></span><br><span data-ttu-id="12b5d-288">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-288">
        - TableBindings</span></span><br><span data-ttu-id="12b5d-289">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-289">
        - TableCoercion</span></span><br><span data-ttu-id="12b5d-290">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-290">
        - TextBindings</span></span><br><span data-ttu-id="12b5d-291">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-291">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-292">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="12b5d-292">Office 2019 for Mac</span></span><br><span data-ttu-id="12b5d-293">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-293">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="12b5d-294">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-294">- TaskPane</span></span><br><span data-ttu-id="12b5d-295">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-295">
        - Content</span></span><br><span data-ttu-id="12b5d-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="12b5d-297">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-297">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12b5d-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12b5d-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12b5d-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12b5d-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12b5d-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12b5d-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12b5d-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12b5d-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="12b5d-307">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-307">- BindingEvents</span></span><br><span data-ttu-id="12b5d-308">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-308">
        - CompressedFile</span></span><br><span data-ttu-id="12b5d-309">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-309">
        - DocumentEvents</span></span><br><span data-ttu-id="12b5d-310">
        - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-310">
        - File</span></span><br><span data-ttu-id="12b5d-311">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-311">
        - MatrixBindings</span></span><br><span data-ttu-id="12b5d-312">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-312">
        - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-313">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-313">
        - PdfFile</span></span><br><span data-ttu-id="12b5d-314">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-314">
        - Selection</span></span><br><span data-ttu-id="12b5d-315">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-315">
        - Settings</span></span><br><span data-ttu-id="12b5d-316">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-316">
        - TableBindings</span></span><br><span data-ttu-id="12b5d-317">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-317">
        - TableCoercion</span></span><br><span data-ttu-id="12b5d-318">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-318">
        - TextBindings</span></span><br><span data-ttu-id="12b5d-319">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-319">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-320">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="12b5d-320">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="12b5d-321">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-321">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="12b5d-322">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-322">- TaskPane</span></span><br><span data-ttu-id="12b5d-323">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-323">
        - Content</span></span></td>
    <td><span data-ttu-id="12b5d-324">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-324">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12b5d-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="12b5d-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="12b5d-326">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-326">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="12b5d-327">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-327">- BindingEvents</span></span><br><span data-ttu-id="12b5d-328">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-328">
        - CompressedFile</span></span><br><span data-ttu-id="12b5d-329">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-329">
        - DocumentEvents</span></span><br><span data-ttu-id="12b5d-330">
        - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-330">
        - File</span></span><br><span data-ttu-id="12b5d-331">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-331">
        - MatrixBindings</span></span><br><span data-ttu-id="12b5d-332">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-332">
        - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-333">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-333">
        - PdfFile</span></span><br><span data-ttu-id="12b5d-334">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-334">
        - Selection</span></span><br><span data-ttu-id="12b5d-335">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-335">
        - Settings</span></span><br><span data-ttu-id="12b5d-336">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-336">
        - TableBindings</span></span><br><span data-ttu-id="12b5d-337">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-337">
        - TableCoercion</span></span><br><span data-ttu-id="12b5d-338">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-338">
        - TextBindings</span></span><br><span data-ttu-id="12b5d-339">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-339">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="12b5d-340">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="12b5d-340">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="12b5d-341">自定义函数</span><span class="sxs-lookup"><span data-stu-id="12b5d-341">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="12b5d-342">平台</span><span class="sxs-lookup"><span data-stu-id="12b5d-342">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="12b5d-343">扩展点</span><span class="sxs-lookup"><span data-stu-id="12b5d-343">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="12b5d-344">API 要求集</span><span class="sxs-lookup"><span data-stu-id="12b5d-344">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="12b5d-345"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="12b5d-345"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-346">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="12b5d-346">Office on the web</span></span></td>
    <td><span data-ttu-id="12b5d-347">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="12b5d-347">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="12b5d-348">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-348">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-349">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="12b5d-349">Office on Windows</span></span><br><span data-ttu-id="12b5d-350">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="12b5d-350">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="12b5d-351">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="12b5d-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="12b5d-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-353">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="12b5d-353">Office for Mac</span></span><br><span data-ttu-id="12b5d-354">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="12b5d-354">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="12b5d-355">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="12b5d-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="12b5d-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="12b5d-357">Outlook</span><span class="sxs-lookup"><span data-stu-id="12b5d-357">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="12b5d-358">平台</span><span class="sxs-lookup"><span data-stu-id="12b5d-358">Platform</span></span></th>
    <th><span data-ttu-id="12b5d-359">扩展点</span><span class="sxs-lookup"><span data-stu-id="12b5d-359">Extension points</span></span></th>
    <th><span data-ttu-id="12b5d-360">API 要求集</span><span class="sxs-lookup"><span data-stu-id="12b5d-360">API requirement sets</span></span></th>
    <th><span data-ttu-id="12b5d-361"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="12b5d-361"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-362">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="12b5d-362">Office on the web</span></span><br><span data-ttu-id="12b5d-363">（新式）</span><span class="sxs-lookup"><span data-stu-id="12b5d-363">Modern</span></span></td>
    <td> <span data-ttu-id="12b5d-364">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="12b5d-364">- Mail Read</span></span><br><span data-ttu-id="12b5d-365">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="12b5d-365">
      - Mail Compose</span></span><br><span data-ttu-id="12b5d-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12b5d-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12b5d-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12b5d-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12b5d-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12b5d-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="12b5d-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="12b5d-374">不可用</span><span class="sxs-lookup"><span data-stu-id="12b5d-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-375">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="12b5d-375">Office on the web</span></span><br><span data-ttu-id="12b5d-376">（经典）</span><span class="sxs-lookup"><span data-stu-id="12b5d-376">Classic.</span></span></td>
    <td> <span data-ttu-id="12b5d-377">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="12b5d-377">- Mail Read</span></span><br><span data-ttu-id="12b5d-378">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="12b5d-378">
      - Mail Compose</span></span><br><span data-ttu-id="12b5d-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12b5d-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12b5d-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12b5d-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12b5d-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12b5d-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="12b5d-386">不可用</span><span class="sxs-lookup"><span data-stu-id="12b5d-386">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-387">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="12b5d-387">Office on Windows</span></span><br><span data-ttu-id="12b5d-388">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="12b5d-388">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="12b5d-389">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="12b5d-389">- Mail Read</span></span><br><span data-ttu-id="12b5d-390">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="12b5d-390">
      - Mail Compose</span></span><br><span data-ttu-id="12b5d-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="12b5d-392">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="12b5d-392">
      - Modules</span></span></td>
    <td> <span data-ttu-id="12b5d-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12b5d-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12b5d-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12b5d-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12b5d-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12b5d-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="12b5d-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="12b5d-400">不可用</span><span class="sxs-lookup"><span data-stu-id="12b5d-400">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-401">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="12b5d-401">Office 2019 on Windows</span></span><br><span data-ttu-id="12b5d-402">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-402">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-403">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="12b5d-403">- Mail Read</span></span><br><span data-ttu-id="12b5d-404">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="12b5d-404">
      - Mail Compose</span></span><br><span data-ttu-id="12b5d-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="12b5d-406">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="12b5d-406">
      - Modules</span></span></td>
    <td> <span data-ttu-id="12b5d-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12b5d-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12b5d-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12b5d-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12b5d-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12b5d-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="12b5d-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="12b5d-414">不可用</span><span class="sxs-lookup"><span data-stu-id="12b5d-414">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-415">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="12b5d-415">Office 2016 on Windows</span></span><br><span data-ttu-id="12b5d-416">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-416">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-417">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="12b5d-417">- Mail Read</span></span><br><span data-ttu-id="12b5d-418">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="12b5d-418">
      - Mail Compose</span></span><br><span data-ttu-id="12b5d-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="12b5d-420">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="12b5d-420">
      - Modules</span></span></td>
    <td> <span data-ttu-id="12b5d-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12b5d-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12b5d-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12b5d-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="12b5d-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="12b5d-425">不可用</span><span class="sxs-lookup"><span data-stu-id="12b5d-425">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-426">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="12b5d-426">Office 2013 on Windows</span></span><br><span data-ttu-id="12b5d-427">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-427">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-428">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="12b5d-428">- Mail Read</span></span><br><span data-ttu-id="12b5d-429">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="12b5d-429">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="12b5d-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12b5d-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12b5d-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="12b5d-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="12b5d-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="12b5d-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="12b5d-434">不可用</span><span class="sxs-lookup"><span data-stu-id="12b5d-434">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-435">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="12b5d-435">Office apps on iOS</span></span><br><span data-ttu-id="12b5d-436">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="12b5d-436">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="12b5d-437">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="12b5d-437">- Mail Read</span></span><br><span data-ttu-id="12b5d-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-439">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-439">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12b5d-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12b5d-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12b5d-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12b5d-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="12b5d-444">不可用</span><span class="sxs-lookup"><span data-stu-id="12b5d-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-445">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="12b5d-445">Office apps on Mac</span></span><br><span data-ttu-id="12b5d-446">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="12b5d-446">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="12b5d-447">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="12b5d-447">- Mail Read</span></span><br><span data-ttu-id="12b5d-448">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="12b5d-448">
      - Mail Compose</span></span><br><span data-ttu-id="12b5d-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12b5d-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12b5d-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12b5d-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12b5d-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12b5d-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="12b5d-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="12b5d-457">不可用</span><span class="sxs-lookup"><span data-stu-id="12b5d-457">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-458">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="12b5d-458">Office 2019 for Mac</span></span><br><span data-ttu-id="12b5d-459">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-459">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-460">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="12b5d-460">- Mail Read</span></span><br><span data-ttu-id="12b5d-461">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="12b5d-461">
      - Mail Compose</span></span><br><span data-ttu-id="12b5d-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12b5d-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12b5d-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12b5d-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12b5d-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12b5d-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="12b5d-469">不可用</span><span class="sxs-lookup"><span data-stu-id="12b5d-469">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-470">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="12b5d-470">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="12b5d-471">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-471">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-472">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="12b5d-472">- Mail Read</span></span><br><span data-ttu-id="12b5d-473">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="12b5d-473">
      - Mail Compose</span></span><br><span data-ttu-id="12b5d-474">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-474">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-475">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-475">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12b5d-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12b5d-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12b5d-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12b5d-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12b5d-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="12b5d-481">不可用</span><span class="sxs-lookup"><span data-stu-id="12b5d-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-482">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="12b5d-482">Office apps on Android</span></span><br><span data-ttu-id="12b5d-483">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="12b5d-483">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="12b5d-484">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="12b5d-484">- Mail Read</span></span><br><span data-ttu-id="12b5d-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12b5d-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12b5d-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12b5d-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12b5d-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="12b5d-491">不可用</span><span class="sxs-lookup"><span data-stu-id="12b5d-491">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="12b5d-492">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="12b5d-492">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="12b5d-493">Word</span><span class="sxs-lookup"><span data-stu-id="12b5d-493">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="12b5d-494">平台</span><span class="sxs-lookup"><span data-stu-id="12b5d-494">Platform</span></span></th>
    <th><span data-ttu-id="12b5d-495">扩展点</span><span class="sxs-lookup"><span data-stu-id="12b5d-495">Extension points</span></span></th>
    <th><span data-ttu-id="12b5d-496">API 要求集</span><span class="sxs-lookup"><span data-stu-id="12b5d-496">API requirement sets</span></span></th>
    <th><span data-ttu-id="12b5d-497"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="12b5d-497"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-498">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="12b5d-498">Office on the web</span></span></td>
    <td> <span data-ttu-id="12b5d-499">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-499">- TaskPane</span></span><br><span data-ttu-id="12b5d-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-501">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-501">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="12b5d-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="12b5d-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="12b5d-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="12b5d-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="12b5d-507">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-507">- BindingEvents</span></span><br><span data-ttu-id="12b5d-508">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12b5d-508">
         - CustomXmlParts</span></span><br><span data-ttu-id="12b5d-509">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-509">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-510">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-510">
         - File</span></span><br><span data-ttu-id="12b5d-511">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-511">
         - HtmlCoercion</span></span><br><span data-ttu-id="12b5d-512">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-512">
         - MatrixBindings</span></span><br><span data-ttu-id="12b5d-513">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-513">
         - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-514">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-514">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12b5d-515">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-515">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-516">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-516">
         - Selection</span></span><br><span data-ttu-id="12b5d-517">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-517">
         - Settings</span></span><br><span data-ttu-id="12b5d-518">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-518">
         - TableBindings</span></span><br><span data-ttu-id="12b5d-519">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-519">
         - TableCoercion</span></span><br><span data-ttu-id="12b5d-520">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-520">
         - TextBindings</span></span><br><span data-ttu-id="12b5d-521">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-521">
         - TextCoercion</span></span><br><span data-ttu-id="12b5d-522">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-522">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-523">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="12b5d-523">Office on Windows</span></span><br><span data-ttu-id="12b5d-524">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="12b5d-524">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="12b5d-525">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-525">- TaskPane</span></span><br><span data-ttu-id="12b5d-526">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-526">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-527">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-527">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="12b5d-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="12b5d-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="12b5d-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="12b5d-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="12b5d-533">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-533">- BindingEvents</span></span><br><span data-ttu-id="12b5d-534">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-534">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-535">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12b5d-535">
         - CustomXmlParts</span></span><br><span data-ttu-id="12b5d-536">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-536">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-537">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-537">
         - File</span></span><br><span data-ttu-id="12b5d-538">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-538">
         - HtmlCoercion</span></span><br><span data-ttu-id="12b5d-539">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-539">
         - MatrixBindings</span></span><br><span data-ttu-id="12b5d-540">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-540">
         - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-541">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-541">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12b5d-542">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-542">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-543">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-543">
         - Selection</span></span><br><span data-ttu-id="12b5d-544">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-544">
         - Settings</span></span><br><span data-ttu-id="12b5d-545">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-545">
         - TableBindings</span></span><br><span data-ttu-id="12b5d-546">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-546">
         - TableCoercion</span></span><br><span data-ttu-id="12b5d-547">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-547">
         - TextBindings</span></span><br><span data-ttu-id="12b5d-548">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-548">
         - TextCoercion</span></span><br><span data-ttu-id="12b5d-549">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-549">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-550">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="12b5d-550">Office 2019 on Windows</span></span><br><span data-ttu-id="12b5d-551">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-551">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-552">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-552">- TaskPane</span></span><br><span data-ttu-id="12b5d-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="12b5d-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="12b5d-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="12b5d-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="12b5d-559">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-559">- BindingEvents</span></span><br><span data-ttu-id="12b5d-560">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-560">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-561">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12b5d-561">
         - CustomXmlParts</span></span><br><span data-ttu-id="12b5d-562">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-562">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-563">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-563">
         - File</span></span><br><span data-ttu-id="12b5d-564">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-564">
         - HtmlCoercion</span></span><br><span data-ttu-id="12b5d-565">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-565">
         - MatrixBindings</span></span><br><span data-ttu-id="12b5d-566">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-566">
         - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-567">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-567">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12b5d-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-568">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-569">
         - Selection</span></span><br><span data-ttu-id="12b5d-570">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-570">
         - Settings</span></span><br><span data-ttu-id="12b5d-571">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-571">
         - TableBindings</span></span><br><span data-ttu-id="12b5d-572">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-572">
         - TableCoercion</span></span><br><span data-ttu-id="12b5d-573">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-573">
         - TextBindings</span></span><br><span data-ttu-id="12b5d-574">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-574">
         - TextCoercion</span></span><br><span data-ttu-id="12b5d-575">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-575">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-576">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="12b5d-576">Office 2016 on Windows</span></span><br><span data-ttu-id="12b5d-577">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-577">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-578">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-578">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12b5d-579">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-579">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="12b5d-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="12b5d-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="12b5d-581">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-581">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="12b5d-582">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-582">- BindingEvents</span></span><br><span data-ttu-id="12b5d-583">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-583">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-584">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12b5d-584">
         - CustomXmlParts</span></span><br><span data-ttu-id="12b5d-585">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-585">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-586">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-586">
         - File</span></span><br><span data-ttu-id="12b5d-587">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-587">
         - HtmlCoercion</span></span><br><span data-ttu-id="12b5d-588">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-588">
         - MatrixBindings</span></span><br><span data-ttu-id="12b5d-589">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-589">
         - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-590">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-590">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12b5d-591">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-591">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-592">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-592">
         - Selection</span></span><br><span data-ttu-id="12b5d-593">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-593">
         - Settings</span></span><br><span data-ttu-id="12b5d-594">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-594">
         - TableBindings</span></span><br><span data-ttu-id="12b5d-595">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-595">
         - TableCoercion</span></span><br><span data-ttu-id="12b5d-596">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-596">
         - TextBindings</span></span><br><span data-ttu-id="12b5d-597">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-597">
         - TextCoercion</span></span><br><span data-ttu-id="12b5d-598">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-598">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-599">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="12b5d-599">Office 2013 on Windows</span></span><br><span data-ttu-id="12b5d-600">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-600">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-601">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-601">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12b5d-602">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="12b5d-602">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="12b5d-603">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-603">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="12b5d-604">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-604">- BindingEvents</span></span><br><span data-ttu-id="12b5d-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-605">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-606">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12b5d-606">
         - CustomXmlParts</span></span><br><span data-ttu-id="12b5d-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-607">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-608">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-608">
         - File</span></span><br><span data-ttu-id="12b5d-609">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-609">
         - HtmlCoercion</span></span><br><span data-ttu-id="12b5d-610">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-610">
         - MatrixBindings</span></span><br><span data-ttu-id="12b5d-611">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-611">
         - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-612">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-612">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12b5d-613">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-613">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-614">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-614">
         - Selection</span></span><br><span data-ttu-id="12b5d-615">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-615">
         - Settings</span></span><br><span data-ttu-id="12b5d-616">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-616">
         - TableBindings</span></span><br><span data-ttu-id="12b5d-617">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-617">
         - TableCoercion</span></span><br><span data-ttu-id="12b5d-618">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-618">
         - TextBindings</span></span><br><span data-ttu-id="12b5d-619">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-619">
         - TextCoercion</span></span><br><span data-ttu-id="12b5d-620">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-620">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-621">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="12b5d-621">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="12b5d-622">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="12b5d-622">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="12b5d-623">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-623">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12b5d-624">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-624">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="12b5d-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="12b5d-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="12b5d-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="12b5d-629">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-629">- BindingEvents</span></span><br><span data-ttu-id="12b5d-630">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-630">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-631">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12b5d-631">
         - CustomXmlParts</span></span><br><span data-ttu-id="12b5d-632">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-632">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-633">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-633">
         - File</span></span><br><span data-ttu-id="12b5d-634">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-634">
         - HtmlCoercion</span></span><br><span data-ttu-id="12b5d-635">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-635">
         - MatrixBindings</span></span><br><span data-ttu-id="12b5d-636">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-636">
         - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-637">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-637">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12b5d-638">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-638">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-639">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-639">
         - Selection</span></span><br><span data-ttu-id="12b5d-640">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-640">
         - Settings</span></span><br><span data-ttu-id="12b5d-641">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-641">
         - TableBindings</span></span><br><span data-ttu-id="12b5d-642">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-642">
         - TableCoercion</span></span><br><span data-ttu-id="12b5d-643">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-643">
         - TextBindings</span></span><br><span data-ttu-id="12b5d-644">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-644">
         - TextCoercion</span></span><br><span data-ttu-id="12b5d-645">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-645">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-646">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="12b5d-646">Office apps on Mac</span></span><br><span data-ttu-id="12b5d-647">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="12b5d-647">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="12b5d-648">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-648">- TaskPane</span></span><br><span data-ttu-id="12b5d-649">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-649">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-650">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-650">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="12b5d-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="12b5d-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="12b5d-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="12b5d-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="12b5d-656">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-656">- BindingEvents</span></span><br><span data-ttu-id="12b5d-657">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-657">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-658">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12b5d-658">
         - CustomXmlParts</span></span><br><span data-ttu-id="12b5d-659">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-659">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-660">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-660">
         - File</span></span><br><span data-ttu-id="12b5d-661">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-661">
         - HtmlCoercion</span></span><br><span data-ttu-id="12b5d-662">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-662">
         - MatrixBindings</span></span><br><span data-ttu-id="12b5d-663">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-663">
         - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-664">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-664">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12b5d-665">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-665">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-666">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-666">
         - Selection</span></span><br><span data-ttu-id="12b5d-667">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-667">
         - Settings</span></span><br><span data-ttu-id="12b5d-668">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-668">
         - TableBindings</span></span><br><span data-ttu-id="12b5d-669">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-669">
         - TableCoercion</span></span><br><span data-ttu-id="12b5d-670">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-670">
         - TextBindings</span></span><br><span data-ttu-id="12b5d-671">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-671">
         - TextCoercion</span></span><br><span data-ttu-id="12b5d-672">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-672">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-673">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="12b5d-673">Office 2019 for Mac</span></span><br><span data-ttu-id="12b5d-674">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-674">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-675">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-675">- TaskPane</span></span><br><span data-ttu-id="12b5d-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="12b5d-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="12b5d-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="12b5d-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="12b5d-682">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-682">- BindingEvents</span></span><br><span data-ttu-id="12b5d-683">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-683">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-684">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12b5d-684">
         - CustomXmlParts</span></span><br><span data-ttu-id="12b5d-685">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-685">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-686">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-686">
         - File</span></span><br><span data-ttu-id="12b5d-687">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-687">
         - HtmlCoercion</span></span><br><span data-ttu-id="12b5d-688">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-688">
         - MatrixBindings</span></span><br><span data-ttu-id="12b5d-689">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-689">
         - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-690">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-690">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12b5d-691">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-691">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-692">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-692">
         - Selection</span></span><br><span data-ttu-id="12b5d-693">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-693">
         - Settings</span></span><br><span data-ttu-id="12b5d-694">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-694">
         - TableBindings</span></span><br><span data-ttu-id="12b5d-695">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-695">
         - TableCoercion</span></span><br><span data-ttu-id="12b5d-696">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-696">
         - TextBindings</span></span><br><span data-ttu-id="12b5d-697">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-697">
         - TextCoercion</span></span><br><span data-ttu-id="12b5d-698">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-698">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-699">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="12b5d-699">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="12b5d-700">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-700">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-701">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-701">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12b5d-702">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-702">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="12b5d-703">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="12b5d-703">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="12b5d-704">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-704">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="12b5d-705">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-705">- BindingEvents</span></span><br><span data-ttu-id="12b5d-706">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-706">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-707">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12b5d-707">
         - CustomXmlParts</span></span><br><span data-ttu-id="12b5d-708">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-708">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-709">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-709">
         - File</span></span><br><span data-ttu-id="12b5d-710">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-710">
         - HtmlCoercion</span></span><br><span data-ttu-id="12b5d-711">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-711">
         - MatrixBindings</span></span><br><span data-ttu-id="12b5d-712">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-712">
         - MatrixCoercion</span></span><br><span data-ttu-id="12b5d-713">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-713">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12b5d-714">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-714">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-715">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-715">
         - Selection</span></span><br><span data-ttu-id="12b5d-716">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-716">
         - Settings</span></span><br><span data-ttu-id="12b5d-717">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-717">
         - TableBindings</span></span><br><span data-ttu-id="12b5d-718">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-718">
         - TableCoercion</span></span><br><span data-ttu-id="12b5d-719">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12b5d-719">
         - TextBindings</span></span><br><span data-ttu-id="12b5d-720">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-720">
         - TextCoercion</span></span><br><span data-ttu-id="12b5d-721">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-721">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="12b5d-722">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="12b5d-722">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="12b5d-723">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="12b5d-723">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="12b5d-724">平台</span><span class="sxs-lookup"><span data-stu-id="12b5d-724">Platform</span></span></th>
    <th><span data-ttu-id="12b5d-725">扩展点</span><span class="sxs-lookup"><span data-stu-id="12b5d-725">Extension points</span></span></th>
    <th><span data-ttu-id="12b5d-726">API 要求集</span><span class="sxs-lookup"><span data-stu-id="12b5d-726">API requirement sets</span></span></th>
    <th><span data-ttu-id="12b5d-727"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="12b5d-727"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-728">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="12b5d-728">Office on the web</span></span></td>
    <td> <span data-ttu-id="12b5d-729">- 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-729">- Content</span></span><br><span data-ttu-id="12b5d-730">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-730">
         - TaskPane</span></span><br><span data-ttu-id="12b5d-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-732">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-732">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="12b5d-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="12b5d-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="12b5d-736">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12b5d-736">- ActiveView</span></span><br><span data-ttu-id="12b5d-737">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-737">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-738">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-738">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-739">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-739">
         - File</span></span><br><span data-ttu-id="12b5d-740">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-740">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-741">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-741">
         - Selection</span></span><br><span data-ttu-id="12b5d-742">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-742">
         - Settings</span></span><br><span data-ttu-id="12b5d-743">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-743">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-744">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="12b5d-744">Office on Windows</span></span><br><span data-ttu-id="12b5d-745">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="12b5d-745">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="12b5d-746">- 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-746">- Content</span></span><br><span data-ttu-id="12b5d-747">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-747">
         - TaskPane</span></span><br><span data-ttu-id="12b5d-748">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-748">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-749">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-749">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="12b5d-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="12b5d-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="12b5d-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12b5d-753">- ActiveView</span></span><br><span data-ttu-id="12b5d-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-754">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-755">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-756">
         - File</span></span><br><span data-ttu-id="12b5d-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-757">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-758">
         - Selection</span></span><br><span data-ttu-id="12b5d-759">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-759">
         - Settings</span></span><br><span data-ttu-id="12b5d-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-761">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="12b5d-761">Office 2019 on Windows</span></span><br><span data-ttu-id="12b5d-762">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-763">- 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-763">- Content</span></span><br><span data-ttu-id="12b5d-764">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-764">
         - TaskPane</span></span><br><span data-ttu-id="12b5d-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="12b5d-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12b5d-768">- ActiveView</span></span><br><span data-ttu-id="12b5d-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-769">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-770">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-771">
         - File</span></span><br><span data-ttu-id="12b5d-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-772">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-773">
         - Selection</span></span><br><span data-ttu-id="12b5d-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-774">
         - Settings</span></span><br><span data-ttu-id="12b5d-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-776">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="12b5d-776">Office 2016 on Windows</span></span><br><span data-ttu-id="12b5d-777">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-778">- 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-778">- Content</span></span><br><span data-ttu-id="12b5d-779">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="12b5d-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="12b5d-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="12b5d-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="12b5d-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12b5d-782">- ActiveView</span></span><br><span data-ttu-id="12b5d-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-783">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-784">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-785">
         - File</span></span><br><span data-ttu-id="12b5d-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-786">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-787">
         - Selection</span></span><br><span data-ttu-id="12b5d-788">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-788">
         - Settings</span></span><br><span data-ttu-id="12b5d-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-790">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="12b5d-790">Office 2013 on Windows</span></span><br><span data-ttu-id="12b5d-791">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-792">- 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-792">- Content</span></span><br><span data-ttu-id="12b5d-793">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="12b5d-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="12b5d-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="12b5d-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="12b5d-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12b5d-796">- ActiveView</span></span><br><span data-ttu-id="12b5d-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-797">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-798">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-799">
         - File</span></span><br><span data-ttu-id="12b5d-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-800">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-801">
         - Selection</span></span><br><span data-ttu-id="12b5d-802">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-802">
         - Settings</span></span><br><span data-ttu-id="12b5d-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-804">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="12b5d-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="12b5d-805">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="12b5d-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="12b5d-806">- 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-806">- Content</span></span><br><span data-ttu-id="12b5d-807">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="12b5d-808">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-808">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="12b5d-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="12b5d-811">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12b5d-811">- ActiveView</span></span><br><span data-ttu-id="12b5d-812">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-812">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-813">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-813">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-814">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-814">
         - File</span></span><br><span data-ttu-id="12b5d-815">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-815">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-816">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-816">
         - Selection</span></span><br><span data-ttu-id="12b5d-817">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-817">
         - Settings</span></span><br><span data-ttu-id="12b5d-818">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-818">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-819">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="12b5d-819">Office apps on Mac</span></span><br><span data-ttu-id="12b5d-820">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="12b5d-820">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="12b5d-821">- 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-821">- Content</span></span><br><span data-ttu-id="12b5d-822">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-822">
         - TaskPane</span></span><br><span data-ttu-id="12b5d-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-824">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-824">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="12b5d-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="12b5d-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="12b5d-828">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12b5d-828">- ActiveView</span></span><br><span data-ttu-id="12b5d-829">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-829">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-830">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-830">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-831">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-831">
         - File</span></span><br><span data-ttu-id="12b5d-832">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-832">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-833">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-833">
         - Selection</span></span><br><span data-ttu-id="12b5d-834">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-834">
         - Settings</span></span><br><span data-ttu-id="12b5d-835">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-835">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-836">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="12b5d-836">Office 2019 for Mac</span></span><br><span data-ttu-id="12b5d-837">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-837">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-838">- 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-838">- Content</span></span><br><span data-ttu-id="12b5d-839">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-839">
         - TaskPane</span></span><br><span data-ttu-id="12b5d-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-841">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-841">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="12b5d-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12b5d-843">- ActiveView</span></span><br><span data-ttu-id="12b5d-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-844">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-845">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-846">
         - File</span></span><br><span data-ttu-id="12b5d-847">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-847">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-848">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-848">
         - Selection</span></span><br><span data-ttu-id="12b5d-849">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-849">
         - Settings</span></span><br><span data-ttu-id="12b5d-850">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-850">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-851">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="12b5d-851">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="12b5d-852">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-852">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-853">- 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-853">- Content</span></span><br><span data-ttu-id="12b5d-854">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-854">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="12b5d-855">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="12b5d-855">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="12b5d-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="12b5d-857">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12b5d-857">- ActiveView</span></span><br><span data-ttu-id="12b5d-858">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-858">
         - CompressedFile</span></span><br><span data-ttu-id="12b5d-859">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-859">
         - DocumentEvents</span></span><br><span data-ttu-id="12b5d-860">
         - File</span><span class="sxs-lookup"><span data-stu-id="12b5d-860">
         - File</span></span><br><span data-ttu-id="12b5d-861">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12b5d-861">
         - PdfFile</span></span><br><span data-ttu-id="12b5d-862">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-862">
         - Selection</span></span><br><span data-ttu-id="12b5d-863">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-863">
         - Settings</span></span><br><span data-ttu-id="12b5d-864">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-864">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="12b5d-865">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="12b5d-865">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="12b5d-866">OneNote</span><span class="sxs-lookup"><span data-stu-id="12b5d-866">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="12b5d-867">平台</span><span class="sxs-lookup"><span data-stu-id="12b5d-867">Platform</span></span></th>
    <th><span data-ttu-id="12b5d-868">扩展点</span><span class="sxs-lookup"><span data-stu-id="12b5d-868">Extension points</span></span></th>
    <th><span data-ttu-id="12b5d-869">API 要求集</span><span class="sxs-lookup"><span data-stu-id="12b5d-869">API requirement sets</span></span></th>
    <th><span data-ttu-id="12b5d-870"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="12b5d-870"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-871">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="12b5d-871">Office on the web</span></span></td>
    <td> <span data-ttu-id="12b5d-872">- 内容</span><span class="sxs-lookup"><span data-stu-id="12b5d-872">- Content</span></span><br><span data-ttu-id="12b5d-873">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-873">
         - TaskPane</span></span><br><span data-ttu-id="12b5d-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12b5d-875">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-875">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="12b5d-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="12b5d-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="12b5d-878">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12b5d-878">- DocumentEvents</span></span><br><span data-ttu-id="12b5d-879">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-879">
         - HtmlCoercion</span></span><br><span data-ttu-id="12b5d-880">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="12b5d-880">
         - Settings</span></span><br><span data-ttu-id="12b5d-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-881">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="12b5d-882">项目</span><span class="sxs-lookup"><span data-stu-id="12b5d-882">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="12b5d-883">平台</span><span class="sxs-lookup"><span data-stu-id="12b5d-883">Platform</span></span></th>
    <th><span data-ttu-id="12b5d-884">扩展点</span><span class="sxs-lookup"><span data-stu-id="12b5d-884">Extension points</span></span></th>
    <th><span data-ttu-id="12b5d-885">API 要求集</span><span class="sxs-lookup"><span data-stu-id="12b5d-885">API requirement sets</span></span></th>
    <th><span data-ttu-id="12b5d-886"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="12b5d-886"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-887">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="12b5d-887">Office 2019 on Windows</span></span><br><span data-ttu-id="12b5d-888">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-888">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-889">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-889">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12b5d-890">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-890">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12b5d-891">- Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-891">- Selection</span></span><br><span data-ttu-id="12b5d-892">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-892">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-893">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="12b5d-893">Office 2016 on Windows</span></span><br><span data-ttu-id="12b5d-894">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-894">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-895">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-895">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12b5d-896">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-896">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12b5d-897">- Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-897">- Selection</span></span><br><span data-ttu-id="12b5d-898">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-898">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12b5d-899">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="12b5d-899">Office 2013 on Windows</span></span><br><span data-ttu-id="12b5d-900">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="12b5d-900">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12b5d-901">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="12b5d-901">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12b5d-902">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12b5d-902">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12b5d-903">- Selection</span><span class="sxs-lookup"><span data-stu-id="12b5d-903">- Selection</span></span><br><span data-ttu-id="12b5d-904">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12b5d-904">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="12b5d-905">另请参阅</span><span class="sxs-lookup"><span data-stu-id="12b5d-905">See also</span></span>

- [<span data-ttu-id="12b5d-906">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="12b5d-906">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="12b5d-907">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="12b5d-907">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="12b5d-908">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="12b5d-908">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="12b5d-909">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="12b5d-909">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="12b5d-910">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="12b5d-910">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="12b5d-911">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="12b5d-911">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="12b5d-912">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="12b5d-912">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="12b5d-913">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="12b5d-913">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="12b5d-914">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="12b5d-914">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="12b5d-915">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="12b5d-915">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="12b5d-916">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="12b5d-916">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)

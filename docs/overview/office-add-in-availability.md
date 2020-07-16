---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 06/23/2020
localization_priority: Priority
ms.openlocfilehash: 979c873b1c5f2d1d7847414f037d5c75737aa33d
ms.sourcegitcommit: a4873c3525c7d30ef551545d27eb2c0a16b4eb50
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/25/2020
ms.locfileid: "44888157"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="b8a95-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="b8a95-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="b8a95-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="b8a95-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="b8a95-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="b8a95-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="b8a95-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="b8a95-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="b8a95-108">Excel</span><span class="sxs-lookup"><span data-stu-id="b8a95-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b8a95-109">平台</span><span class="sxs-lookup"><span data-stu-id="b8a95-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b8a95-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="b8a95-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b8a95-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b8a95-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b8a95-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b8a95-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="b8a95-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="b8a95-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-114">- TaskPane</span></span><br><span data-ttu-id="b8a95-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-115">
        - Content</span></span><br><span data-ttu-id="b8a95-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="b8a95-116">
        - Custom Functions</span></span><br><span data-ttu-id="b8a95-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="b8a95-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b8a95-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8a95-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8a95-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8a95-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8a95-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8a95-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8a95-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b8a95-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b8a95-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b8a95-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b8a95-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="b8a95-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="b8a95-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b8a95-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-131">
        - BindingEvents</span></span><br><span data-ttu-id="b8a95-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-132">
        - CompressedFile</span></span><br><span data-ttu-id="b8a95-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-133">
        - DocumentEvents</span></span><br><span data-ttu-id="b8a95-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-134">
        - File</span></span><br><span data-ttu-id="b8a95-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-135">
        - MatrixBindings</span></span><br><span data-ttu-id="b8a95-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-137">
        - Selection</span></span><br><span data-ttu-id="b8a95-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-138">
        - Settings</span></span><br><span data-ttu-id="b8a95-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-139">
        - TableBindings</span></span><br><span data-ttu-id="b8a95-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-140">
        - TableCoercion</span></span><br><span data-ttu-id="b8a95-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-141">
        - TextBindings</span></span><br><span data-ttu-id="b8a95-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-143">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="b8a95-143">Office on Windows</span></span><br><span data-ttu-id="b8a95-144">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="b8a95-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8a95-145">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-145">- TaskPane</span></span><br><span data-ttu-id="b8a95-146">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-146">
        - Content</span></span><br><span data-ttu-id="b8a95-147">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="b8a95-147">
        - Custom Functions</span></span><br><span data-ttu-id="b8a95-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="b8a95-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b8a95-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8a95-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8a95-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8a95-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8a95-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8a95-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8a95-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b8a95-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b8a95-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b8a95-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b8a95-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="b8a95-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8a95-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="b8a95-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-163">
        - BindingEvents</span></span><br><span data-ttu-id="b8a95-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-164">
        - CompressedFile</span></span><br><span data-ttu-id="b8a95-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-165">
        - DocumentEvents</span></span><br><span data-ttu-id="b8a95-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-166">
        - File</span></span><br><span data-ttu-id="b8a95-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-167">
        - MatrixBindings</span></span><br><span data-ttu-id="b8a95-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-169">
        - Selection</span></span><br><span data-ttu-id="b8a95-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-170">
        - Settings</span></span><br><span data-ttu-id="b8a95-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-171">
        - TableBindings</span></span><br><span data-ttu-id="b8a95-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-172">
        - TableCoercion</span></span><br><span data-ttu-id="b8a95-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-173">
        - TextBindings</span></span><br><span data-ttu-id="b8a95-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-175">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="b8a95-175">Office 2019 on Windows</span></span><br><span data-ttu-id="b8a95-176">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b8a95-177">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-177">- TaskPane</span></span><br><span data-ttu-id="b8a95-178">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-178">
        - Content</span></span><br><span data-ttu-id="b8a95-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b8a95-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8a95-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8a95-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8a95-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8a95-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8a95-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8a95-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b8a95-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b8a95-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b8a95-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-190">- BindingEvents</span></span><br><span data-ttu-id="b8a95-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-191">
        - CompressedFile</span></span><br><span data-ttu-id="b8a95-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-192">
        - DocumentEvents</span></span><br><span data-ttu-id="b8a95-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-193">
        - File</span></span><br><span data-ttu-id="b8a95-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-194">
        - MatrixBindings</span></span><br><span data-ttu-id="b8a95-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-196">
        - Selection</span></span><br><span data-ttu-id="b8a95-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-197">
        - Settings</span></span><br><span data-ttu-id="b8a95-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-198">
        - TableBindings</span></span><br><span data-ttu-id="b8a95-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-199">
        - TableCoercion</span></span><br><span data-ttu-id="b8a95-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-200">
        - TextBindings</span></span><br><span data-ttu-id="b8a95-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-202">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="b8a95-202">Office 2016 on Windows</span></span><br><span data-ttu-id="b8a95-203">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b8a95-204">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-204">- TaskPane</span></span><br><span data-ttu-id="b8a95-205">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-205">
        - Content</span></span></td>
    <td><span data-ttu-id="b8a95-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8a95-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b8a95-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b8a95-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b8a95-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-209">- BindingEvents</span></span><br><span data-ttu-id="b8a95-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-210">
        - CompressedFile</span></span><br><span data-ttu-id="b8a95-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-211">
        - DocumentEvents</span></span><br><span data-ttu-id="b8a95-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-212">
        - File</span></span><br><span data-ttu-id="b8a95-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-213">
        - MatrixBindings</span></span><br><span data-ttu-id="b8a95-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-215">
        - Selection</span></span><br><span data-ttu-id="b8a95-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-216">
        - Settings</span></span><br><span data-ttu-id="b8a95-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-217">
        - TableBindings</span></span><br><span data-ttu-id="b8a95-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-218">
        - TableCoercion</span></span><br><span data-ttu-id="b8a95-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-219">
        - TextBindings</span></span><br><span data-ttu-id="b8a95-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-221">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="b8a95-221">Office 2013 on Windows</span></span><br><span data-ttu-id="b8a95-222">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b8a95-223">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-223">
        - TaskPane</span></span><br><span data-ttu-id="b8a95-224">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="b8a95-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b8a95-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b8a95-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b8a95-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-227">
        - BindingEvents</span></span><br><span data-ttu-id="b8a95-228">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-228">
        - DocumentEvents</span></span><br><span data-ttu-id="b8a95-229">
        - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-229">
        - File</span></span><br><span data-ttu-id="b8a95-230">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-230">
        - MatrixBindings</span></span><br><span data-ttu-id="b8a95-231">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-231">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-232">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-232">
        - Selection</span></span><br><span data-ttu-id="b8a95-233">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-233">
        - Settings</span></span><br><span data-ttu-id="b8a95-234">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-234">
        - TableBindings</span></span><br><span data-ttu-id="b8a95-235">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-235">
        - TableCoercion</span></span><br><span data-ttu-id="b8a95-236">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-236">
        - TextBindings</span></span><br><span data-ttu-id="b8a95-237">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-237">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-238">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="b8a95-238">Office on iPad</span></span><br><span data-ttu-id="b8a95-239">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="b8a95-239">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b8a95-240">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-240">- TaskPane</span></span><br><span data-ttu-id="b8a95-241">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-241">
        - Content</span></span></td>
    <td><span data-ttu-id="b8a95-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8a95-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8a95-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8a95-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8a95-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8a95-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8a95-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b8a95-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b8a95-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b8a95-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b8a95-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="b8a95-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b8a95-255">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-255">- BindingEvents</span></span><br><span data-ttu-id="b8a95-256">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-256">
        - DocumentEvents</span></span><br><span data-ttu-id="b8a95-257">
        - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-257">
        - File</span></span><br><span data-ttu-id="b8a95-258">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-258">
        - MatrixBindings</span></span><br><span data-ttu-id="b8a95-259">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-259">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-260">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-260">
        - Selection</span></span><br><span data-ttu-id="b8a95-261">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-261">
        - Settings</span></span><br><span data-ttu-id="b8a95-262">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-262">
        - TableBindings</span></span><br><span data-ttu-id="b8a95-263">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-263">
        - TableCoercion</span></span><br><span data-ttu-id="b8a95-264">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-264">
        - TextBindings</span></span><br><span data-ttu-id="b8a95-265">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-265">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-266">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="b8a95-266">Office on Mac</span></span><br><span data-ttu-id="b8a95-267">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="b8a95-267">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b8a95-268">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-268">- TaskPane</span></span><br><span data-ttu-id="b8a95-269">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-269">
        - Content</span></span><br><span data-ttu-id="b8a95-270">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="b8a95-270">
        - Custom Functions</span></span><br><span data-ttu-id="b8a95-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b8a95-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8a95-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8a95-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8a95-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8a95-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8a95-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8a95-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b8a95-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b8a95-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b8a95-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b8a95-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="b8a95-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8a95-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="b8a95-286">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-286">- BindingEvents</span></span><br><span data-ttu-id="b8a95-287">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-287">
        - CompressedFile</span></span><br><span data-ttu-id="b8a95-288">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-288">
        - DocumentEvents</span></span><br><span data-ttu-id="b8a95-289">
        - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-289">
        - File</span></span><br><span data-ttu-id="b8a95-290">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-290">
        - MatrixBindings</span></span><br><span data-ttu-id="b8a95-291">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-291">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-292">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-292">
        - PdfFile</span></span><br><span data-ttu-id="b8a95-293">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-293">
        - Selection</span></span><br><span data-ttu-id="b8a95-294">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-294">
        - Settings</span></span><br><span data-ttu-id="b8a95-295">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-295">
        - TableBindings</span></span><br><span data-ttu-id="b8a95-296">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-296">
        - TableCoercion</span></span><br><span data-ttu-id="b8a95-297">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-297">
        - TextBindings</span></span><br><span data-ttu-id="b8a95-298">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-298">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-299">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="b8a95-299">Office 2019 on Mac</span></span><br><span data-ttu-id="b8a95-300">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-300">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b8a95-301">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-301">- TaskPane</span></span><br><span data-ttu-id="b8a95-302">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-302">
        - Content</span></span><br><span data-ttu-id="b8a95-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b8a95-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8a95-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8a95-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8a95-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8a95-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8a95-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8a95-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b8a95-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b8a95-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b8a95-314">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-314">- BindingEvents</span></span><br><span data-ttu-id="b8a95-315">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-315">
        - CompressedFile</span></span><br><span data-ttu-id="b8a95-316">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-316">
        - DocumentEvents</span></span><br><span data-ttu-id="b8a95-317">
        - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-317">
        - File</span></span><br><span data-ttu-id="b8a95-318">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-318">
        - MatrixBindings</span></span><br><span data-ttu-id="b8a95-319">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-319">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-320">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-320">
        - PdfFile</span></span><br><span data-ttu-id="b8a95-321">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-321">
        - Selection</span></span><br><span data-ttu-id="b8a95-322">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-322">
        - Settings</span></span><br><span data-ttu-id="b8a95-323">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-323">
        - TableBindings</span></span><br><span data-ttu-id="b8a95-324">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-324">
        - TableCoercion</span></span><br><span data-ttu-id="b8a95-325">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-325">
        - TextBindings</span></span><br><span data-ttu-id="b8a95-326">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-326">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-327">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="b8a95-327">Office 2016 on Mac</span></span><br><span data-ttu-id="b8a95-328">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b8a95-329">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-329">- TaskPane</span></span><br><span data-ttu-id="b8a95-330">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-330">
        - Content</span></span></td>
    <td><span data-ttu-id="b8a95-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8a95-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b8a95-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b8a95-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b8a95-334">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-334">- BindingEvents</span></span><br><span data-ttu-id="b8a95-335">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-335">
        - CompressedFile</span></span><br><span data-ttu-id="b8a95-336">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-336">
        - DocumentEvents</span></span><br><span data-ttu-id="b8a95-337">
        - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-337">
        - File</span></span><br><span data-ttu-id="b8a95-338">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-338">
        - MatrixBindings</span></span><br><span data-ttu-id="b8a95-339">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-339">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-340">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-340">
        - PdfFile</span></span><br><span data-ttu-id="b8a95-341">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-341">
        - Selection</span></span><br><span data-ttu-id="b8a95-342">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-342">
        - Settings</span></span><br><span data-ttu-id="b8a95-343">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-343">
        - TableBindings</span></span><br><span data-ttu-id="b8a95-344">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-344">
        - TableCoercion</span></span><br><span data-ttu-id="b8a95-345">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-345">
        - TextBindings</span></span><br><span data-ttu-id="b8a95-346">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-346">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b8a95-347">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="b8a95-347">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="b8a95-348">自定义函数（仅 Excel）</span><span class="sxs-lookup"><span data-stu-id="b8a95-348">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b8a95-349">平台</span><span class="sxs-lookup"><span data-stu-id="b8a95-349">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b8a95-350">扩展点</span><span class="sxs-lookup"><span data-stu-id="b8a95-350">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b8a95-351">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b8a95-351">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b8a95-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b8a95-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-353">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="b8a95-353">Office on the web</span></span></td>
    <td><span data-ttu-id="b8a95-354">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="b8a95-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b8a95-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-356">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="b8a95-356">Office on Windows</span></span><br><span data-ttu-id="b8a95-357">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="b8a95-357">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b8a95-358">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="b8a95-358">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b8a95-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-360">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="b8a95-360">Office on Mac</span></span><br><span data-ttu-id="b8a95-361">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="b8a95-361">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="b8a95-362">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="b8a95-362">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b8a95-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="b8a95-364">Outlook</span><span class="sxs-lookup"><span data-stu-id="b8a95-364">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b8a95-365">平台</span><span class="sxs-lookup"><span data-stu-id="b8a95-365">Platform</span></span></th>
    <th><span data-ttu-id="b8a95-366">扩展点</span><span class="sxs-lookup"><span data-stu-id="b8a95-366">Extension points</span></span></th>
    <th><span data-ttu-id="b8a95-367">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b8a95-367">API requirement sets</span></span></th>
    <th><span data-ttu-id="b8a95-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b8a95-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-369">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="b8a95-369">Office on the web</span></span><br><span data-ttu-id="b8a95-370">（新式）</span><span class="sxs-lookup"><span data-stu-id="b8a95-370">(modern)</span></span></td>
    <td> <span data-ttu-id="b8a95-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b8a95-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b8a95-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b8a95-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b8a95-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8a95-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8a95-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8a95-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8a95-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8a95-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b8a95-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b8a95-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b8a95-384">不可用</span><span class="sxs-lookup"><span data-stu-id="b8a95-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-385">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="b8a95-385">Office on the web</span></span><br><span data-ttu-id="b8a95-386">（经典）</span><span class="sxs-lookup"><span data-stu-id="b8a95-386">(classic)</span></span></td>
    <td> <span data-ttu-id="b8a95-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b8a95-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b8a95-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b8a95-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b8a95-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8a95-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8a95-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8a95-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8a95-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8a95-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b8a95-398">不可用</span><span class="sxs-lookup"><span data-stu-id="b8a95-398">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-399">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="b8a95-399">Office on Windows</span></span><br><span data-ttu-id="b8a95-400">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="b8a95-400">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8a95-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b8a95-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b8a95-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b8a95-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b8a95-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b8a95-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="b8a95-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8a95-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8a95-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8a95-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8a95-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8a95-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b8a95-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b8a95-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b8a95-415">不可用</span><span class="sxs-lookup"><span data-stu-id="b8a95-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-416">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="b8a95-416">Office 2019 on Windows</span></span><br><span data-ttu-id="b8a95-417">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-417">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b8a95-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b8a95-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b8a95-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b8a95-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b8a95-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="b8a95-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8a95-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8a95-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8a95-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8a95-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8a95-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b8a95-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b8a95-431">不可用</span><span class="sxs-lookup"><span data-stu-id="b8a95-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-432">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="b8a95-432">Office 2016 on Windows</span></span><br><span data-ttu-id="b8a95-433">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b8a95-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b8a95-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b8a95-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b8a95-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b8a95-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="b8a95-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8a95-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8a95-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8a95-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="b8a95-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="b8a95-444">不可用</span><span class="sxs-lookup"><span data-stu-id="b8a95-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-445">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="b8a95-445">Office 2013 on Windows</span></span><br><span data-ttu-id="b8a95-446">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-446">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b8a95-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b8a95-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b8a95-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="b8a95-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8a95-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8a95-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="b8a95-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="b8a95-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="b8a95-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="b8a95-455">不可用</span><span class="sxs-lookup"><span data-stu-id="b8a95-455">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-456">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="b8a95-456">Office on iOS</span></span><br><span data-ttu-id="b8a95-457">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="b8a95-457">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8a95-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b8a95-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8a95-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8a95-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8a95-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8a95-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b8a95-465">不可用</span><span class="sxs-lookup"><span data-stu-id="b8a95-465">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-466">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="b8a95-466">Office on Mac</span></span><br><span data-ttu-id="b8a95-467">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="b8a95-467">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8a95-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b8a95-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b8a95-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b8a95-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b8a95-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8a95-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8a95-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8a95-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8a95-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8a95-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b8a95-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b8a95-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b8a95-481">不可用</span><span class="sxs-lookup"><span data-stu-id="b8a95-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-482">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="b8a95-482">Office 2019 on Mac</span></span><br><span data-ttu-id="b8a95-483">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-483">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b8a95-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b8a95-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b8a95-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b8a95-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8a95-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8a95-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8a95-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8a95-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8a95-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b8a95-495">不可用</span><span class="sxs-lookup"><span data-stu-id="b8a95-495">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-496">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="b8a95-496">Office 2016 on Mac</span></span><br><span data-ttu-id="b8a95-497">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-497">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b8a95-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b8a95-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b8a95-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b8a95-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8a95-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8a95-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8a95-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8a95-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8a95-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b8a95-509">不可用</span><span class="sxs-lookup"><span data-stu-id="b8a95-509">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-510">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="b8a95-510">Office on Android</span></span><br><span data-ttu-id="b8a95-511">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="b8a95-511">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8a95-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b8a95-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">约会组织者（撰写）：联机会议</a> （预览）</span><span class="sxs-lookup"><span data-stu-id="b8a95-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="b8a95-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8a95-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8a95-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8a95-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8a95-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b8a95-520">不可用</span><span class="sxs-lookup"><span data-stu-id="b8a95-520">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="b8a95-521">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="b8a95-521">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b8a95-522">要求集的客户端支持可能受到 Exchange 服务器支持的限制。</span><span class="sxs-lookup"><span data-stu-id="b8a95-522">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="b8a95-523">有关 Exchange 服务器和 Outlook 客户端支持的要求集范围的详细信息，请参阅 [Outlook JavaScript API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="b8a95-523">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="b8a95-524">Word</span><span class="sxs-lookup"><span data-stu-id="b8a95-524">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b8a95-525">平台</span><span class="sxs-lookup"><span data-stu-id="b8a95-525">Platform</span></span></th>
    <th><span data-ttu-id="b8a95-526">扩展点</span><span class="sxs-lookup"><span data-stu-id="b8a95-526">Extension points</span></span></th>
    <th><span data-ttu-id="b8a95-527">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b8a95-527">API requirement sets</span></span></th>
    <th><span data-ttu-id="b8a95-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b8a95-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-529">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="b8a95-529">Office on the web</span></span></td>
    <td> <span data-ttu-id="b8a95-530">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-530">- TaskPane</span></span><br><span data-ttu-id="b8a95-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8a95-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b8a95-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b8a95-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8a95-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b8a95-538">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-538">- BindingEvents</span></span><br><span data-ttu-id="b8a95-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8a95-539">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8a95-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-540">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-541">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-541">
         - File</span></span><br><span data-ttu-id="b8a95-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-542">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8a95-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-543">
         - MatrixBindings</span></span><br><span data-ttu-id="b8a95-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-544">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-545">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8a95-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-546">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-547">
         - Selection</span></span><br><span data-ttu-id="b8a95-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-548">
         - Settings</span></span><br><span data-ttu-id="b8a95-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-549">
         - TableBindings</span></span><br><span data-ttu-id="b8a95-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-550">
         - TableCoercion</span></span><br><span data-ttu-id="b8a95-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-551">
         - TextBindings</span></span><br><span data-ttu-id="b8a95-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-552">
         - TextCoercion</span></span><br><span data-ttu-id="b8a95-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-553">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-554">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="b8a95-554">Office on Windows</span></span><br><span data-ttu-id="b8a95-555">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="b8a95-555">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8a95-556">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-556">- TaskPane</span></span><br><span data-ttu-id="b8a95-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8a95-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b8a95-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b8a95-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8a95-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b8a95-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-564">- BindingEvents</span></span><br><span data-ttu-id="b8a95-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-565">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8a95-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8a95-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-567">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-568">
         - File</span></span><br><span data-ttu-id="b8a95-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8a95-570">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-570">
         - MatrixBindings</span></span><br><span data-ttu-id="b8a95-571">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-571">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-572">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-572">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8a95-573">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-573">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-574">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-574">
         - Selection</span></span><br><span data-ttu-id="b8a95-575">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-575">
         - Settings</span></span><br><span data-ttu-id="b8a95-576">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-576">
         - TableBindings</span></span><br><span data-ttu-id="b8a95-577">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-577">
         - TableCoercion</span></span><br><span data-ttu-id="b8a95-578">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-578">
         - TextBindings</span></span><br><span data-ttu-id="b8a95-579">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-579">
         - TextCoercion</span></span><br><span data-ttu-id="b8a95-580">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-580">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-581">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="b8a95-581">Office 2019 on Windows</span></span><br><span data-ttu-id="b8a95-582">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-582">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-583">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-583">- TaskPane</span></span><br><span data-ttu-id="b8a95-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8a95-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b8a95-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b8a95-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8a95-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-590">- BindingEvents</span></span><br><span data-ttu-id="b8a95-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-591">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8a95-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8a95-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-593">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-594">
         - File</span></span><br><span data-ttu-id="b8a95-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8a95-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-596">
         - MatrixBindings</span></span><br><span data-ttu-id="b8a95-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8a95-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-599">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-600">
         - Selection</span></span><br><span data-ttu-id="b8a95-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-601">
         - Settings</span></span><br><span data-ttu-id="b8a95-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-602">
         - TableBindings</span></span><br><span data-ttu-id="b8a95-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-603">
         - TableCoercion</span></span><br><span data-ttu-id="b8a95-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-604">
         - TextBindings</span></span><br><span data-ttu-id="b8a95-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-605">
         - TextCoercion</span></span><br><span data-ttu-id="b8a95-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-607">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="b8a95-607">Office 2016 on Windows</span></span><br><span data-ttu-id="b8a95-608">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-608">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-609">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-609">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b8a95-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8a95-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b8a95-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b8a95-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8a95-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-613">- BindingEvents</span></span><br><span data-ttu-id="b8a95-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-614">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8a95-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8a95-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-616">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-617">
         - File</span></span><br><span data-ttu-id="b8a95-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8a95-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-619">
         - MatrixBindings</span></span><br><span data-ttu-id="b8a95-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8a95-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-622">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-623">
         - Selection</span></span><br><span data-ttu-id="b8a95-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-624">
         - Settings</span></span><br><span data-ttu-id="b8a95-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-625">
         - TableBindings</span></span><br><span data-ttu-id="b8a95-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-626">
         - TableCoercion</span></span><br><span data-ttu-id="b8a95-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-627">
         - TextBindings</span></span><br><span data-ttu-id="b8a95-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-628">
         - TextCoercion</span></span><br><span data-ttu-id="b8a95-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-629">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-630">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="b8a95-630">Office 2013 on Windows</span></span><br><span data-ttu-id="b8a95-631">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-631">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-632">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b8a95-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b8a95-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b8a95-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8a95-635">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-635">- BindingEvents</span></span><br><span data-ttu-id="b8a95-636">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-636">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-637">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8a95-637">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8a95-638">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-638">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-639">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-639">
         - File</span></span><br><span data-ttu-id="b8a95-640">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-640">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8a95-641">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-641">
         - MatrixBindings</span></span><br><span data-ttu-id="b8a95-642">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-642">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-643">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-643">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8a95-644">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-644">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-645">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-645">
         - Selection</span></span><br><span data-ttu-id="b8a95-646">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-646">
         - Settings</span></span><br><span data-ttu-id="b8a95-647">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-647">
         - TableBindings</span></span><br><span data-ttu-id="b8a95-648">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-648">
         - TableCoercion</span></span><br><span data-ttu-id="b8a95-649">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-649">
         - TextBindings</span></span><br><span data-ttu-id="b8a95-650">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-650">
         - TextCoercion</span></span><br><span data-ttu-id="b8a95-651">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-651">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-652">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="b8a95-652">Office on iPad</span></span><br><span data-ttu-id="b8a95-653">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="b8a95-653">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8a95-654">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-654">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b8a95-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8a95-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b8a95-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b8a95-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="b8a95-660">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-660">- BindingEvents</span></span><br><span data-ttu-id="b8a95-661">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-661">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-662">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8a95-662">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8a95-663">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-663">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-664">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-664">
         - File</span></span><br><span data-ttu-id="b8a95-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8a95-666">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-666">
         - MatrixBindings</span></span><br><span data-ttu-id="b8a95-667">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-667">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-668">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-668">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8a95-669">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-669">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-670">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-670">
         - Selection</span></span><br><span data-ttu-id="b8a95-671">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-671">
         - Settings</span></span><br><span data-ttu-id="b8a95-672">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-672">
         - TableBindings</span></span><br><span data-ttu-id="b8a95-673">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-673">
         - TableCoercion</span></span><br><span data-ttu-id="b8a95-674">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-674">
         - TextBindings</span></span><br><span data-ttu-id="b8a95-675">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-675">
         - TextCoercion</span></span><br><span data-ttu-id="b8a95-676">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-676">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-677">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="b8a95-677">Office on Mac</span></span><br><span data-ttu-id="b8a95-678">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="b8a95-678">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8a95-679">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-679">- TaskPane</span></span><br><span data-ttu-id="b8a95-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8a95-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b8a95-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b8a95-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8a95-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="b8a95-687">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-687">- BindingEvents</span></span><br><span data-ttu-id="b8a95-688">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-688">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-689">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8a95-689">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8a95-690">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-690">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-691">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-691">
         - File</span></span><br><span data-ttu-id="b8a95-692">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-692">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8a95-693">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-693">
         - MatrixBindings</span></span><br><span data-ttu-id="b8a95-694">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-694">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-695">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-695">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8a95-696">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-696">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-697">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-697">
         - Selection</span></span><br><span data-ttu-id="b8a95-698">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-698">
         - Settings</span></span><br><span data-ttu-id="b8a95-699">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-699">
         - TableBindings</span></span><br><span data-ttu-id="b8a95-700">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-700">
         - TableCoercion</span></span><br><span data-ttu-id="b8a95-701">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-701">
         - TextBindings</span></span><br><span data-ttu-id="b8a95-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-702">
         - TextCoercion</span></span><br><span data-ttu-id="b8a95-703">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-703">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-704">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="b8a95-704">Office 2019 on Mac</span></span><br><span data-ttu-id="b8a95-705">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-705">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-706">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-706">- TaskPane</span></span><br><span data-ttu-id="b8a95-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8a95-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b8a95-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b8a95-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="b8a95-713">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-713">- BindingEvents</span></span><br><span data-ttu-id="b8a95-714">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-714">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-715">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8a95-715">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8a95-716">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-716">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-717">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-717">
         - File</span></span><br><span data-ttu-id="b8a95-718">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-718">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8a95-719">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-719">
         - MatrixBindings</span></span><br><span data-ttu-id="b8a95-720">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-720">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-721">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-721">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8a95-722">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-722">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-723">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-723">
         - Selection</span></span><br><span data-ttu-id="b8a95-724">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-724">
         - Settings</span></span><br><span data-ttu-id="b8a95-725">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-725">
         - TableBindings</span></span><br><span data-ttu-id="b8a95-726">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-726">
         - TableCoercion</span></span><br><span data-ttu-id="b8a95-727">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-727">
         - TextBindings</span></span><br><span data-ttu-id="b8a95-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-728">
         - TextCoercion</span></span><br><span data-ttu-id="b8a95-729">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-729">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-730">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="b8a95-730">Office 2016 on Mac</span></span><br><span data-ttu-id="b8a95-731">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-731">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-732">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-732">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b8a95-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8a95-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b8a95-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b8a95-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8a95-736">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-736">- BindingEvents</span></span><br><span data-ttu-id="b8a95-737">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-737">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-738">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8a95-738">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8a95-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-739">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-740">
         - File</span></span><br><span data-ttu-id="b8a95-741">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-741">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8a95-742">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-742">
         - MatrixBindings</span></span><br><span data-ttu-id="b8a95-743">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-743">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8a95-744">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-744">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8a95-745">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-745">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-746">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-746">
         - Selection</span></span><br><span data-ttu-id="b8a95-747">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-747">
         - Settings</span></span><br><span data-ttu-id="b8a95-748">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-748">
         - TableBindings</span></span><br><span data-ttu-id="b8a95-749">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-749">
         - TableCoercion</span></span><br><span data-ttu-id="b8a95-750">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8a95-750">
         - TextBindings</span></span><br><span data-ttu-id="b8a95-751">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-751">
         - TextCoercion</span></span><br><span data-ttu-id="b8a95-752">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-752">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="b8a95-753">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="b8a95-753">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="b8a95-754">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b8a95-754">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b8a95-755">平台</span><span class="sxs-lookup"><span data-stu-id="b8a95-755">Platform</span></span></th>
    <th><span data-ttu-id="b8a95-756">扩展点</span><span class="sxs-lookup"><span data-stu-id="b8a95-756">Extension points</span></span></th>
    <th><span data-ttu-id="b8a95-757">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b8a95-757">API requirement sets</span></span></th>
    <th><span data-ttu-id="b8a95-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b8a95-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-759">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="b8a95-759">Office on the web</span></span></td>
    <td> <span data-ttu-id="b8a95-760">- 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-760">- Content</span></span><br><span data-ttu-id="b8a95-761">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-761">
         - TaskPane</span></span><br><span data-ttu-id="b8a95-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b8a95-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8a95-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b8a95-767">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8a95-767">- ActiveView</span></span><br><span data-ttu-id="b8a95-768">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-768">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-769">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-769">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-770">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-770">
         - File</span></span><br><span data-ttu-id="b8a95-771">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-771">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-772">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-772">
         - Selection</span></span><br><span data-ttu-id="b8a95-773">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-773">
         - Settings</span></span><br><span data-ttu-id="b8a95-774">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-774">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-775">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="b8a95-775">Office on Windows</span></span><br><span data-ttu-id="b8a95-776">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="b8a95-776">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8a95-777">- 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-777">- Content</span></span><br><span data-ttu-id="b8a95-778">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-778">
         - TaskPane</span></span><br><span data-ttu-id="b8a95-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b8a95-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8a95-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b8a95-784">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8a95-784">- ActiveView</span></span><br><span data-ttu-id="b8a95-785">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-785">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-786">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-786">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-787">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-787">
         - File</span></span><br><span data-ttu-id="b8a95-788">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-788">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-789">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-789">
         - Selection</span></span><br><span data-ttu-id="b8a95-790">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-790">
         - Settings</span></span><br><span data-ttu-id="b8a95-791">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-791">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-792">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="b8a95-792">Office 2019 on Windows</span></span><br><span data-ttu-id="b8a95-793">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-793">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-794">- 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-794">- Content</span></span><br><span data-ttu-id="b8a95-795">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-795">
         - TaskPane</span></span><br><span data-ttu-id="b8a95-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8a95-799">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8a95-799">- ActiveView</span></span><br><span data-ttu-id="b8a95-800">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-800">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-801">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-801">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-802">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-802">
         - File</span></span><br><span data-ttu-id="b8a95-803">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-803">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-804">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-804">
         - Selection</span></span><br><span data-ttu-id="b8a95-805">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-805">
         - Settings</span></span><br><span data-ttu-id="b8a95-806">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-806">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-807">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="b8a95-807">Office 2016 on Windows</span></span><br><span data-ttu-id="b8a95-808">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-808">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-809">- 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-809">- Content</span></span><br><span data-ttu-id="b8a95-810">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-810">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b8a95-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b8a95-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b8a95-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8a95-813">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8a95-813">- ActiveView</span></span><br><span data-ttu-id="b8a95-814">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-814">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-815">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-815">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-816">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-816">
         - File</span></span><br><span data-ttu-id="b8a95-817">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-817">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-818">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-818">
         - Selection</span></span><br><span data-ttu-id="b8a95-819">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-819">
         - Settings</span></span><br><span data-ttu-id="b8a95-820">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-820">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-821">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="b8a95-821">Office 2013 on Windows</span></span><br><span data-ttu-id="b8a95-822">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-822">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-823">- 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-823">- Content</span></span><br><span data-ttu-id="b8a95-824">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-824">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="b8a95-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b8a95-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b8a95-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8a95-827">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8a95-827">- ActiveView</span></span><br><span data-ttu-id="b8a95-828">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-828">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-829">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-829">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-830">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-830">
         - File</span></span><br><span data-ttu-id="b8a95-831">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-831">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-832">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-832">
         - Selection</span></span><br><span data-ttu-id="b8a95-833">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-833">
         - Settings</span></span><br><span data-ttu-id="b8a95-834">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-834">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-835">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="b8a95-835">Office on iPad</span></span><br><span data-ttu-id="b8a95-836">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="b8a95-836">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8a95-837">- 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-837">- Content</span></span><br><span data-ttu-id="b8a95-838">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-838">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b8a95-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b8a95-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8a95-842">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8a95-842">- ActiveView</span></span><br><span data-ttu-id="b8a95-843">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-843">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-844">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-844">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-845">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-845">
         - File</span></span><br><span data-ttu-id="b8a95-846">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-846">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-847">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-847">
         - Selection</span></span><br><span data-ttu-id="b8a95-848">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-848">
         - Settings</span></span><br><span data-ttu-id="b8a95-849">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-849">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-850">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="b8a95-850">Office on Mac</span></span><br><span data-ttu-id="b8a95-851">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="b8a95-851">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8a95-852">- 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-852">- Content</span></span><br><span data-ttu-id="b8a95-853">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-853">
         - TaskPane</span></span><br><span data-ttu-id="b8a95-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b8a95-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8a95-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b8a95-859">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8a95-859">- ActiveView</span></span><br><span data-ttu-id="b8a95-860">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-860">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-861">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-861">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-862">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-862">
         - File</span></span><br><span data-ttu-id="b8a95-863">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-863">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-864">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-864">
         - Selection</span></span><br><span data-ttu-id="b8a95-865">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-865">
         - Settings</span></span><br><span data-ttu-id="b8a95-866">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-866">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-867">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="b8a95-867">Office 2019 on Mac</span></span><br><span data-ttu-id="b8a95-868">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-868">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-869">- 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-869">- Content</span></span><br><span data-ttu-id="b8a95-870">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-870">
         - TaskPane</span></span><br><span data-ttu-id="b8a95-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8a95-874">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8a95-874">- ActiveView</span></span><br><span data-ttu-id="b8a95-875">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-875">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-876">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-876">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-877">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-877">
         - File</span></span><br><span data-ttu-id="b8a95-878">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-878">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-879">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-879">
         - Selection</span></span><br><span data-ttu-id="b8a95-880">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-880">
         - Settings</span></span><br><span data-ttu-id="b8a95-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-881">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-882">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="b8a95-882">Office 2016 on Mac</span></span><br><span data-ttu-id="b8a95-883">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-883">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-884">- 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-884">- Content</span></span><br><span data-ttu-id="b8a95-885">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-885">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b8a95-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b8a95-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b8a95-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8a95-888">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8a95-888">- ActiveView</span></span><br><span data-ttu-id="b8a95-889">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-889">
         - CompressedFile</span></span><br><span data-ttu-id="b8a95-890">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-890">
         - DocumentEvents</span></span><br><span data-ttu-id="b8a95-891">
         - File</span><span class="sxs-lookup"><span data-stu-id="b8a95-891">
         - File</span></span><br><span data-ttu-id="b8a95-892">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8a95-892">
         - PdfFile</span></span><br><span data-ttu-id="b8a95-893">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-893">
         - Selection</span></span><br><span data-ttu-id="b8a95-894">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-894">
         - Settings</span></span><br><span data-ttu-id="b8a95-895">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-895">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b8a95-896">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="b8a95-896">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="b8a95-897">OneNote</span><span class="sxs-lookup"><span data-stu-id="b8a95-897">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b8a95-898">平台</span><span class="sxs-lookup"><span data-stu-id="b8a95-898">Platform</span></span></th>
    <th><span data-ttu-id="b8a95-899">扩展点</span><span class="sxs-lookup"><span data-stu-id="b8a95-899">Extension points</span></span></th>
    <th><span data-ttu-id="b8a95-900">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b8a95-900">API requirement sets</span></span></th>
    <th><span data-ttu-id="b8a95-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b8a95-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-902">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="b8a95-902">Office on the web</span></span></td>
    <td> <span data-ttu-id="b8a95-903">- 内容</span><span class="sxs-lookup"><span data-stu-id="b8a95-903">- Content</span></span><br><span data-ttu-id="b8a95-904">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-904">
         - TaskPane</span></span><br><span data-ttu-id="b8a95-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8a95-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="b8a95-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8a95-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8a95-909">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8a95-909">- DocumentEvents</span></span><br><span data-ttu-id="b8a95-910">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-910">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8a95-911">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b8a95-911">
         - Settings</span></span><br><span data-ttu-id="b8a95-912">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-912">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="b8a95-913">项目</span><span class="sxs-lookup"><span data-stu-id="b8a95-913">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b8a95-914">平台</span><span class="sxs-lookup"><span data-stu-id="b8a95-914">Platform</span></span></th>
    <th><span data-ttu-id="b8a95-915">扩展点</span><span class="sxs-lookup"><span data-stu-id="b8a95-915">Extension points</span></span></th>
    <th><span data-ttu-id="b8a95-916">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b8a95-916">API requirement sets</span></span></th>
    <th><span data-ttu-id="b8a95-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b8a95-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-918">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="b8a95-918">Office 2019 on Windows</span></span><br><span data-ttu-id="b8a95-919">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-919">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-920">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-920">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b8a95-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8a95-922">- Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-922">- Selection</span></span><br><span data-ttu-id="b8a95-923">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-923">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-924">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="b8a95-924">Office 2016 on Windows</span></span><br><span data-ttu-id="b8a95-925">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-925">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-926">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-926">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b8a95-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8a95-928">- Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-928">- Selection</span></span><br><span data-ttu-id="b8a95-929">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-929">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8a95-930">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="b8a95-930">Office 2013 on Windows</span></span><br><span data-ttu-id="b8a95-931">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="b8a95-931">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8a95-932">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b8a95-932">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b8a95-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8a95-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8a95-934">- Selection</span><span class="sxs-lookup"><span data-stu-id="b8a95-934">- Selection</span></span><br><span data-ttu-id="b8a95-935">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8a95-935">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="b8a95-936">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b8a95-936">See also</span></span>

- [<span data-ttu-id="b8a95-937">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="b8a95-937">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="b8a95-938">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="b8a95-938">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="b8a95-939">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="b8a95-939">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="b8a95-940">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="b8a95-940">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="b8a95-941">API 参考文档</span><span class="sxs-lookup"><span data-stu-id="b8a95-941">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="b8a95-942">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="b8a95-942">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="b8a95-943">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="b8a95-943">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="b8a95-944">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="b8a95-944">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="b8a95-945">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="b8a95-945">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="b8a95-946">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="b8a95-946">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="b8a95-947">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="b8a95-947">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="b8a95-948">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="b8a95-948">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
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
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="d21a2-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="d21a2-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="d21a2-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span><span class="sxs-lookup"><span data-stu-id="d21a2-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="d21a2-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span><span class="sxs-lookup"><span data-stu-id="d21a2-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="d21a2-106">通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="d21a2-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="d21a2-107">有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。</span><span class="sxs-lookup"><span data-stu-id="d21a2-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="d21a2-108">Excel</span><span class="sxs-lookup"><span data-stu-id="d21a2-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="d21a2-109">平台</span><span class="sxs-lookup"><span data-stu-id="d21a2-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="d21a2-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="d21a2-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="d21a2-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="d21a2-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="d21a2-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="d21a2-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="d21a2-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="d21a2-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-114">- TaskPane</span></span><br><span data-ttu-id="d21a2-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-115">
        - Content</span></span><br><span data-ttu-id="d21a2-116">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="d21a2-116">
        - Custom Functions</span></span><br><span data-ttu-id="d21a2-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="d21a2-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="d21a2-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d21a2-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d21a2-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d21a2-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d21a2-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d21a2-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d21a2-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d21a2-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d21a2-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d21a2-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="d21a2-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="d21a2-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="d21a2-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d21a2-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-131">
        - BindingEvents</span></span><br><span data-ttu-id="d21a2-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-132">
        - CompressedFile</span></span><br><span data-ttu-id="d21a2-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-133">
        - DocumentEvents</span></span><br><span data-ttu-id="d21a2-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-134">
        - File</span></span><br><span data-ttu-id="d21a2-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-135">
        - MatrixBindings</span></span><br><span data-ttu-id="d21a2-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-137">
        - Selection</span></span><br><span data-ttu-id="d21a2-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-138">
        - Settings</span></span><br><span data-ttu-id="d21a2-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-139">
        - TableBindings</span></span><br><span data-ttu-id="d21a2-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-140">
        - TableCoercion</span></span><br><span data-ttu-id="d21a2-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-141">
        - TextBindings</span></span><br><span data-ttu-id="d21a2-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-143">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="d21a2-143">Office on Windows</span></span><br><span data-ttu-id="d21a2-144">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d21a2-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d21a2-145">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-145">- TaskPane</span></span><br><span data-ttu-id="d21a2-146">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-146">
        - Content</span></span><br><span data-ttu-id="d21a2-147">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="d21a2-147">
        - Custom Functions</span></span><br><span data-ttu-id="d21a2-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="d21a2-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="d21a2-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d21a2-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d21a2-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d21a2-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d21a2-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d21a2-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d21a2-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d21a2-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d21a2-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d21a2-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="d21a2-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="d21a2-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d21a2-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="d21a2-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-163">
        - BindingEvents</span></span><br><span data-ttu-id="d21a2-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-164">
        - CompressedFile</span></span><br><span data-ttu-id="d21a2-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-165">
        - DocumentEvents</span></span><br><span data-ttu-id="d21a2-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-166">
        - File</span></span><br><span data-ttu-id="d21a2-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-167">
        - MatrixBindings</span></span><br><span data-ttu-id="d21a2-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-169">
        - Selection</span></span><br><span data-ttu-id="d21a2-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-170">
        - Settings</span></span><br><span data-ttu-id="d21a2-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-171">
        - TableBindings</span></span><br><span data-ttu-id="d21a2-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-172">
        - TableCoercion</span></span><br><span data-ttu-id="d21a2-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-173">
        - TextBindings</span></span><br><span data-ttu-id="d21a2-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-175">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="d21a2-175">Office 2019 on Windows</span></span><br><span data-ttu-id="d21a2-176">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d21a2-177">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-177">- TaskPane</span></span><br><span data-ttu-id="d21a2-178">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-178">
        - Content</span></span><br><span data-ttu-id="d21a2-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d21a2-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d21a2-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d21a2-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d21a2-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d21a2-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d21a2-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d21a2-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d21a2-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d21a2-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d21a2-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-190">- BindingEvents</span></span><br><span data-ttu-id="d21a2-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-191">
        - CompressedFile</span></span><br><span data-ttu-id="d21a2-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-192">
        - DocumentEvents</span></span><br><span data-ttu-id="d21a2-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-193">
        - File</span></span><br><span data-ttu-id="d21a2-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-194">
        - MatrixBindings</span></span><br><span data-ttu-id="d21a2-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-196">
        - Selection</span></span><br><span data-ttu-id="d21a2-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-197">
        - Settings</span></span><br><span data-ttu-id="d21a2-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-198">
        - TableBindings</span></span><br><span data-ttu-id="d21a2-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-199">
        - TableCoercion</span></span><br><span data-ttu-id="d21a2-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-200">
        - TextBindings</span></span><br><span data-ttu-id="d21a2-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-202">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="d21a2-202">Office 2016 on Windows</span></span><br><span data-ttu-id="d21a2-203">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d21a2-204">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-204">- TaskPane</span></span><br><span data-ttu-id="d21a2-205">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-205">
        - Content</span></span></td>
    <td><span data-ttu-id="d21a2-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d21a2-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d21a2-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d21a2-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d21a2-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-209">- BindingEvents</span></span><br><span data-ttu-id="d21a2-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-210">
        - CompressedFile</span></span><br><span data-ttu-id="d21a2-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-211">
        - DocumentEvents</span></span><br><span data-ttu-id="d21a2-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-212">
        - File</span></span><br><span data-ttu-id="d21a2-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-213">
        - MatrixBindings</span></span><br><span data-ttu-id="d21a2-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-215">
        - Selection</span></span><br><span data-ttu-id="d21a2-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-216">
        - Settings</span></span><br><span data-ttu-id="d21a2-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-217">
        - TableBindings</span></span><br><span data-ttu-id="d21a2-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-218">
        - TableCoercion</span></span><br><span data-ttu-id="d21a2-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-219">
        - TextBindings</span></span><br><span data-ttu-id="d21a2-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-221">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="d21a2-221">Office 2013 on Windows</span></span><br><span data-ttu-id="d21a2-222">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d21a2-223">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-223">
        - TaskPane</span></span><br><span data-ttu-id="d21a2-224">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="d21a2-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d21a2-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d21a2-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d21a2-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-227">
        - BindingEvents</span></span><br><span data-ttu-id="d21a2-228">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-228">
        - DocumentEvents</span></span><br><span data-ttu-id="d21a2-229">
        - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-229">
        - File</span></span><br><span data-ttu-id="d21a2-230">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-230">
        - MatrixBindings</span></span><br><span data-ttu-id="d21a2-231">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-231">
        - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-232">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-232">
        - Selection</span></span><br><span data-ttu-id="d21a2-233">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-233">
        - Settings</span></span><br><span data-ttu-id="d21a2-234">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-234">
        - TableBindings</span></span><br><span data-ttu-id="d21a2-235">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-235">
        - TableCoercion</span></span><br><span data-ttu-id="d21a2-236">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-236">
        - TextBindings</span></span><br><span data-ttu-id="d21a2-237">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-237">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-238">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="d21a2-238">Office on iPad</span></span><br><span data-ttu-id="d21a2-239">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d21a2-239">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="d21a2-240">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-240">- TaskPane</span></span><br><span data-ttu-id="d21a2-241">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-241">
        - Content</span></span></td>
    <td><span data-ttu-id="d21a2-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d21a2-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d21a2-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d21a2-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d21a2-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d21a2-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d21a2-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d21a2-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d21a2-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d21a2-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="d21a2-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="d21a2-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d21a2-255">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-255">- BindingEvents</span></span><br><span data-ttu-id="d21a2-256">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-256">
        - DocumentEvents</span></span><br><span data-ttu-id="d21a2-257">
        - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-257">
        - File</span></span><br><span data-ttu-id="d21a2-258">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-258">
        - MatrixBindings</span></span><br><span data-ttu-id="d21a2-259">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-259">
        - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-260">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-260">
        - Selection</span></span><br><span data-ttu-id="d21a2-261">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-261">
        - Settings</span></span><br><span data-ttu-id="d21a2-262">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-262">
        - TableBindings</span></span><br><span data-ttu-id="d21a2-263">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-263">
        - TableCoercion</span></span><br><span data-ttu-id="d21a2-264">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-264">
        - TextBindings</span></span><br><span data-ttu-id="d21a2-265">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-265">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-266">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="d21a2-266">Office on Mac</span></span><br><span data-ttu-id="d21a2-267">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d21a2-267">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="d21a2-268">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-268">- TaskPane</span></span><br><span data-ttu-id="d21a2-269">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-269">
        - Content</span></span><br><span data-ttu-id="d21a2-270">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="d21a2-270">
        - Custom Functions</span></span><br><span data-ttu-id="d21a2-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d21a2-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d21a2-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d21a2-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d21a2-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d21a2-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d21a2-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d21a2-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d21a2-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d21a2-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="d21a2-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="d21a2-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="d21a2-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d21a2-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="d21a2-286">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-286">- BindingEvents</span></span><br><span data-ttu-id="d21a2-287">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-287">
        - CompressedFile</span></span><br><span data-ttu-id="d21a2-288">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-288">
        - DocumentEvents</span></span><br><span data-ttu-id="d21a2-289">
        - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-289">
        - File</span></span><br><span data-ttu-id="d21a2-290">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-290">
        - MatrixBindings</span></span><br><span data-ttu-id="d21a2-291">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-291">
        - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-292">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-292">
        - PdfFile</span></span><br><span data-ttu-id="d21a2-293">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-293">
        - Selection</span></span><br><span data-ttu-id="d21a2-294">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-294">
        - Settings</span></span><br><span data-ttu-id="d21a2-295">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-295">
        - TableBindings</span></span><br><span data-ttu-id="d21a2-296">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-296">
        - TableCoercion</span></span><br><span data-ttu-id="d21a2-297">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-297">
        - TextBindings</span></span><br><span data-ttu-id="d21a2-298">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-298">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-299">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="d21a2-299">Office 2019 on Mac</span></span><br><span data-ttu-id="d21a2-300">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-300">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d21a2-301">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-301">- TaskPane</span></span><br><span data-ttu-id="d21a2-302">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-302">
        - Content</span></span><br><span data-ttu-id="d21a2-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d21a2-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d21a2-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d21a2-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d21a2-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d21a2-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d21a2-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d21a2-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d21a2-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d21a2-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d21a2-314">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-314">- BindingEvents</span></span><br><span data-ttu-id="d21a2-315">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-315">
        - CompressedFile</span></span><br><span data-ttu-id="d21a2-316">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-316">
        - DocumentEvents</span></span><br><span data-ttu-id="d21a2-317">
        - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-317">
        - File</span></span><br><span data-ttu-id="d21a2-318">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-318">
        - MatrixBindings</span></span><br><span data-ttu-id="d21a2-319">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-319">
        - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-320">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-320">
        - PdfFile</span></span><br><span data-ttu-id="d21a2-321">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-321">
        - Selection</span></span><br><span data-ttu-id="d21a2-322">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-322">
        - Settings</span></span><br><span data-ttu-id="d21a2-323">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-323">
        - TableBindings</span></span><br><span data-ttu-id="d21a2-324">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-324">
        - TableCoercion</span></span><br><span data-ttu-id="d21a2-325">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-325">
        - TextBindings</span></span><br><span data-ttu-id="d21a2-326">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-326">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-327">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="d21a2-327">Office 2016 on Mac</span></span><br><span data-ttu-id="d21a2-328">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="d21a2-329">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-329">- TaskPane</span></span><br><span data-ttu-id="d21a2-330">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-330">
        - Content</span></span></td>
    <td><span data-ttu-id="d21a2-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d21a2-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d21a2-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d21a2-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="d21a2-334">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-334">- BindingEvents</span></span><br><span data-ttu-id="d21a2-335">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-335">
        - CompressedFile</span></span><br><span data-ttu-id="d21a2-336">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-336">
        - DocumentEvents</span></span><br><span data-ttu-id="d21a2-337">
        - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-337">
        - File</span></span><br><span data-ttu-id="d21a2-338">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-338">
        - MatrixBindings</span></span><br><span data-ttu-id="d21a2-339">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-339">
        - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-340">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-340">
        - PdfFile</span></span><br><span data-ttu-id="d21a2-341">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-341">
        - Selection</span></span><br><span data-ttu-id="d21a2-342">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-342">
        - Settings</span></span><br><span data-ttu-id="d21a2-343">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-343">
        - TableBindings</span></span><br><span data-ttu-id="d21a2-344">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-344">
        - TableCoercion</span></span><br><span data-ttu-id="d21a2-345">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-345">
        - TextBindings</span></span><br><span data-ttu-id="d21a2-346">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-346">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="d21a2-347">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="d21a2-347">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="d21a2-348">自定义函数（仅 Excel）</span><span class="sxs-lookup"><span data-stu-id="d21a2-348">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="d21a2-349">平台</span><span class="sxs-lookup"><span data-stu-id="d21a2-349">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="d21a2-350">扩展点</span><span class="sxs-lookup"><span data-stu-id="d21a2-350">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="d21a2-351">API 要求集</span><span class="sxs-lookup"><span data-stu-id="d21a2-351">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="d21a2-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="d21a2-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-353">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="d21a2-353">Office on the web</span></span></td>
    <td><span data-ttu-id="d21a2-354">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="d21a2-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d21a2-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-356">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="d21a2-356">Office on Windows</span></span><br><span data-ttu-id="d21a2-357">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d21a2-357">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="d21a2-358">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="d21a2-358">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d21a2-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-360">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="d21a2-360">Office on Mac</span></span><br><span data-ttu-id="d21a2-361">（连接到 Office 365）</span><span class="sxs-lookup"><span data-stu-id="d21a2-361">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="d21a2-362">
        - 自定义函数</span><span class="sxs-lookup"><span data-stu-id="d21a2-362">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="d21a2-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="d21a2-364">Outlook</span><span class="sxs-lookup"><span data-stu-id="d21a2-364">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d21a2-365">平台</span><span class="sxs-lookup"><span data-stu-id="d21a2-365">Platform</span></span></th>
    <th><span data-ttu-id="d21a2-366">扩展点</span><span class="sxs-lookup"><span data-stu-id="d21a2-366">Extension points</span></span></th>
    <th><span data-ttu-id="d21a2-367">API 要求集</span><span class="sxs-lookup"><span data-stu-id="d21a2-367">API requirement sets</span></span></th>
    <th><span data-ttu-id="d21a2-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="d21a2-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-369">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="d21a2-369">Office on the web</span></span><br><span data-ttu-id="d21a2-370">（新式）</span><span class="sxs-lookup"><span data-stu-id="d21a2-370">(modern)</span></span></td>
    <td> <span data-ttu-id="d21a2-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d21a2-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d21a2-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d21a2-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d21a2-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d21a2-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d21a2-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d21a2-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d21a2-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d21a2-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d21a2-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="d21a2-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="d21a2-384">不可用</span><span class="sxs-lookup"><span data-stu-id="d21a2-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-385">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="d21a2-385">Office on the web</span></span><br><span data-ttu-id="d21a2-386">（经典）</span><span class="sxs-lookup"><span data-stu-id="d21a2-386">(classic)</span></span></td>
    <td> <span data-ttu-id="d21a2-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d21a2-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d21a2-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d21a2-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d21a2-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d21a2-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d21a2-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d21a2-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d21a2-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d21a2-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d21a2-398">不可用</span><span class="sxs-lookup"><span data-stu-id="d21a2-398">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-399">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="d21a2-399">Office on Windows</span></span><br><span data-ttu-id="d21a2-400">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d21a2-400">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d21a2-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d21a2-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d21a2-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d21a2-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d21a2-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d21a2-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="d21a2-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d21a2-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d21a2-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d21a2-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d21a2-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d21a2-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d21a2-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="d21a2-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="d21a2-415">不可用</span><span class="sxs-lookup"><span data-stu-id="d21a2-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-416">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="d21a2-416">Office 2019 on Windows</span></span><br><span data-ttu-id="d21a2-417">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-417">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d21a2-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d21a2-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d21a2-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d21a2-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d21a2-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="d21a2-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d21a2-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d21a2-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d21a2-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d21a2-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d21a2-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d21a2-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d21a2-431">不可用</span><span class="sxs-lookup"><span data-stu-id="d21a2-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-432">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="d21a2-432">Office 2016 on Windows</span></span><br><span data-ttu-id="d21a2-433">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d21a2-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d21a2-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d21a2-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d21a2-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d21a2-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="d21a2-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d21a2-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d21a2-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d21a2-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="d21a2-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="d21a2-444">不可用</span><span class="sxs-lookup"><span data-stu-id="d21a2-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-445">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="d21a2-445">Office 2013 on Windows</span></span><br><span data-ttu-id="d21a2-446">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-446">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d21a2-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d21a2-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d21a2-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="d21a2-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d21a2-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d21a2-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="d21a2-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="d21a2-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="d21a2-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="d21a2-455">不可用</span><span class="sxs-lookup"><span data-stu-id="d21a2-455">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-456">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="d21a2-456">Office on iOS</span></span><br><span data-ttu-id="d21a2-457">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d21a2-457">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d21a2-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d21a2-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d21a2-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d21a2-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d21a2-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d21a2-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d21a2-465">不可用</span><span class="sxs-lookup"><span data-stu-id="d21a2-465">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-466">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="d21a2-466">Office on Mac</span></span><br><span data-ttu-id="d21a2-467">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d21a2-467">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d21a2-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d21a2-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d21a2-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d21a2-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d21a2-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d21a2-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d21a2-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d21a2-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d21a2-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d21a2-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d21a2-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="d21a2-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="d21a2-481">不可用</span><span class="sxs-lookup"><span data-stu-id="d21a2-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-482">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="d21a2-482">Office 2019 on Mac</span></span><br><span data-ttu-id="d21a2-483">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-483">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d21a2-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d21a2-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d21a2-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d21a2-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d21a2-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d21a2-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d21a2-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d21a2-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d21a2-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d21a2-495">不可用</span><span class="sxs-lookup"><span data-stu-id="d21a2-495">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-496">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="d21a2-496">Office 2016 on Mac</span></span><br><span data-ttu-id="d21a2-497">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-497">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d21a2-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="d21a2-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="d21a2-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="d21a2-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d21a2-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d21a2-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d21a2-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d21a2-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d21a2-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d21a2-509">不可用</span><span class="sxs-lookup"><span data-stu-id="d21a2-509">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-510">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="d21a2-510">Office on Android</span></span><br><span data-ttu-id="d21a2-511">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d21a2-511">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d21a2-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">邮件阅读</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="d21a2-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">约会组织者（撰写）：联机会议</a> （预览）</span><span class="sxs-lookup"><span data-stu-id="d21a2-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="d21a2-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d21a2-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d21a2-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d21a2-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d21a2-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d21a2-520">不可用</span><span class="sxs-lookup"><span data-stu-id="d21a2-520">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="d21a2-521">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="d21a2-521">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d21a2-522">要求集的客户端支持可能受到 Exchange 服务器支持的限制。</span><span class="sxs-lookup"><span data-stu-id="d21a2-522">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="d21a2-523">有关 Exchange 服务器和 Outlook 客户端支持的要求集范围的详细信息，请参阅 [Outlook JavaScript API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="d21a2-523">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="d21a2-524">Word</span><span class="sxs-lookup"><span data-stu-id="d21a2-524">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d21a2-525">平台</span><span class="sxs-lookup"><span data-stu-id="d21a2-525">Platform</span></span></th>
    <th><span data-ttu-id="d21a2-526">扩展点</span><span class="sxs-lookup"><span data-stu-id="d21a2-526">Extension points</span></span></th>
    <th><span data-ttu-id="d21a2-527">API 要求集</span><span class="sxs-lookup"><span data-stu-id="d21a2-527">API requirement sets</span></span></th>
    <th><span data-ttu-id="d21a2-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="d21a2-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-529">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="d21a2-529">Office on the web</span></span></td>
    <td> <span data-ttu-id="d21a2-530">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-530">- TaskPane</span></span><br><span data-ttu-id="d21a2-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d21a2-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d21a2-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d21a2-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d21a2-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d21a2-538">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-538">- BindingEvents</span></span><br><span data-ttu-id="d21a2-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d21a2-539">
         - CustomXmlParts</span></span><br><span data-ttu-id="d21a2-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-540">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-541">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-541">
         - File</span></span><br><span data-ttu-id="d21a2-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-542">
         - HtmlCoercion</span></span><br><span data-ttu-id="d21a2-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-543">
         - MatrixBindings</span></span><br><span data-ttu-id="d21a2-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-544">
         - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-545">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d21a2-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-546">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-547">
         - Selection</span></span><br><span data-ttu-id="d21a2-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-548">
         - Settings</span></span><br><span data-ttu-id="d21a2-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-549">
         - TableBindings</span></span><br><span data-ttu-id="d21a2-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-550">
         - TableCoercion</span></span><br><span data-ttu-id="d21a2-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-551">
         - TextBindings</span></span><br><span data-ttu-id="d21a2-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-552">
         - TextCoercion</span></span><br><span data-ttu-id="d21a2-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-553">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-554">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="d21a2-554">Office on Windows</span></span><br><span data-ttu-id="d21a2-555">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d21a2-555">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d21a2-556">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-556">- TaskPane</span></span><br><span data-ttu-id="d21a2-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d21a2-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d21a2-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d21a2-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d21a2-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d21a2-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-564">- BindingEvents</span></span><br><span data-ttu-id="d21a2-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-565">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d21a2-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="d21a2-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-567">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-568">
         - File</span></span><br><span data-ttu-id="d21a2-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="d21a2-570">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-570">
         - MatrixBindings</span></span><br><span data-ttu-id="d21a2-571">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-571">
         - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-572">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-572">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d21a2-573">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-573">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-574">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-574">
         - Selection</span></span><br><span data-ttu-id="d21a2-575">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-575">
         - Settings</span></span><br><span data-ttu-id="d21a2-576">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-576">
         - TableBindings</span></span><br><span data-ttu-id="d21a2-577">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-577">
         - TableCoercion</span></span><br><span data-ttu-id="d21a2-578">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-578">
         - TextBindings</span></span><br><span data-ttu-id="d21a2-579">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-579">
         - TextCoercion</span></span><br><span data-ttu-id="d21a2-580">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-580">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-581">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="d21a2-581">Office 2019 on Windows</span></span><br><span data-ttu-id="d21a2-582">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-582">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-583">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-583">- TaskPane</span></span><br><span data-ttu-id="d21a2-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d21a2-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d21a2-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d21a2-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d21a2-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-590">- BindingEvents</span></span><br><span data-ttu-id="d21a2-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-591">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d21a2-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="d21a2-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-593">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-594">
         - File</span></span><br><span data-ttu-id="d21a2-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="d21a2-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-596">
         - MatrixBindings</span></span><br><span data-ttu-id="d21a2-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d21a2-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-599">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-600">
         - Selection</span></span><br><span data-ttu-id="d21a2-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-601">
         - Settings</span></span><br><span data-ttu-id="d21a2-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-602">
         - TableBindings</span></span><br><span data-ttu-id="d21a2-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-603">
         - TableCoercion</span></span><br><span data-ttu-id="d21a2-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-604">
         - TextBindings</span></span><br><span data-ttu-id="d21a2-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-605">
         - TextCoercion</span></span><br><span data-ttu-id="d21a2-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-607">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="d21a2-607">Office 2016 on Windows</span></span><br><span data-ttu-id="d21a2-608">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-608">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-609">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-609">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d21a2-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d21a2-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d21a2-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d21a2-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d21a2-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-613">- BindingEvents</span></span><br><span data-ttu-id="d21a2-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-614">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d21a2-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="d21a2-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-616">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-617">
         - File</span></span><br><span data-ttu-id="d21a2-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="d21a2-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-619">
         - MatrixBindings</span></span><br><span data-ttu-id="d21a2-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d21a2-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-622">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-623">
         - Selection</span></span><br><span data-ttu-id="d21a2-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-624">
         - Settings</span></span><br><span data-ttu-id="d21a2-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-625">
         - TableBindings</span></span><br><span data-ttu-id="d21a2-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-626">
         - TableCoercion</span></span><br><span data-ttu-id="d21a2-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-627">
         - TextBindings</span></span><br><span data-ttu-id="d21a2-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-628">
         - TextCoercion</span></span><br><span data-ttu-id="d21a2-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-629">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-630">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="d21a2-630">Office 2013 on Windows</span></span><br><span data-ttu-id="d21a2-631">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-631">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-632">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d21a2-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d21a2-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d21a2-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d21a2-635">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-635">- BindingEvents</span></span><br><span data-ttu-id="d21a2-636">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-636">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-637">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d21a2-637">
         - CustomXmlParts</span></span><br><span data-ttu-id="d21a2-638">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-638">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-639">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-639">
         - File</span></span><br><span data-ttu-id="d21a2-640">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-640">
         - HtmlCoercion</span></span><br><span data-ttu-id="d21a2-641">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-641">
         - MatrixBindings</span></span><br><span data-ttu-id="d21a2-642">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-642">
         - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-643">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-643">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d21a2-644">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-644">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-645">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-645">
         - Selection</span></span><br><span data-ttu-id="d21a2-646">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-646">
         - Settings</span></span><br><span data-ttu-id="d21a2-647">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-647">
         - TableBindings</span></span><br><span data-ttu-id="d21a2-648">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-648">
         - TableCoercion</span></span><br><span data-ttu-id="d21a2-649">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-649">
         - TextBindings</span></span><br><span data-ttu-id="d21a2-650">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-650">
         - TextCoercion</span></span><br><span data-ttu-id="d21a2-651">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-651">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-652">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="d21a2-652">Office on iPad</span></span><br><span data-ttu-id="d21a2-653">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d21a2-653">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d21a2-654">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-654">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d21a2-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d21a2-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d21a2-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d21a2-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="d21a2-660">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-660">- BindingEvents</span></span><br><span data-ttu-id="d21a2-661">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-661">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-662">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d21a2-662">
         - CustomXmlParts</span></span><br><span data-ttu-id="d21a2-663">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-663">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-664">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-664">
         - File</span></span><br><span data-ttu-id="d21a2-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="d21a2-666">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-666">
         - MatrixBindings</span></span><br><span data-ttu-id="d21a2-667">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-667">
         - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-668">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-668">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d21a2-669">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-669">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-670">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-670">
         - Selection</span></span><br><span data-ttu-id="d21a2-671">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-671">
         - Settings</span></span><br><span data-ttu-id="d21a2-672">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-672">
         - TableBindings</span></span><br><span data-ttu-id="d21a2-673">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-673">
         - TableCoercion</span></span><br><span data-ttu-id="d21a2-674">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-674">
         - TextBindings</span></span><br><span data-ttu-id="d21a2-675">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-675">
         - TextCoercion</span></span><br><span data-ttu-id="d21a2-676">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-676">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-677">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="d21a2-677">Office on Mac</span></span><br><span data-ttu-id="d21a2-678">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d21a2-678">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d21a2-679">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-679">- TaskPane</span></span><br><span data-ttu-id="d21a2-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d21a2-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d21a2-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d21a2-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d21a2-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="d21a2-687">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-687">- BindingEvents</span></span><br><span data-ttu-id="d21a2-688">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-688">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-689">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d21a2-689">
         - CustomXmlParts</span></span><br><span data-ttu-id="d21a2-690">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-690">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-691">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-691">
         - File</span></span><br><span data-ttu-id="d21a2-692">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-692">
         - HtmlCoercion</span></span><br><span data-ttu-id="d21a2-693">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-693">
         - MatrixBindings</span></span><br><span data-ttu-id="d21a2-694">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-694">
         - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-695">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-695">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d21a2-696">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-696">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-697">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-697">
         - Selection</span></span><br><span data-ttu-id="d21a2-698">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-698">
         - Settings</span></span><br><span data-ttu-id="d21a2-699">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-699">
         - TableBindings</span></span><br><span data-ttu-id="d21a2-700">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-700">
         - TableCoercion</span></span><br><span data-ttu-id="d21a2-701">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-701">
         - TextBindings</span></span><br><span data-ttu-id="d21a2-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-702">
         - TextCoercion</span></span><br><span data-ttu-id="d21a2-703">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-703">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-704">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="d21a2-704">Office 2019 on Mac</span></span><br><span data-ttu-id="d21a2-705">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-705">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-706">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-706">- TaskPane</span></span><br><span data-ttu-id="d21a2-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d21a2-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="d21a2-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="d21a2-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="d21a2-713">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-713">- BindingEvents</span></span><br><span data-ttu-id="d21a2-714">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-714">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-715">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d21a2-715">
         - CustomXmlParts</span></span><br><span data-ttu-id="d21a2-716">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-716">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-717">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-717">
         - File</span></span><br><span data-ttu-id="d21a2-718">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-718">
         - HtmlCoercion</span></span><br><span data-ttu-id="d21a2-719">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-719">
         - MatrixBindings</span></span><br><span data-ttu-id="d21a2-720">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-720">
         - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-721">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-721">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d21a2-722">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-722">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-723">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-723">
         - Selection</span></span><br><span data-ttu-id="d21a2-724">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-724">
         - Settings</span></span><br><span data-ttu-id="d21a2-725">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-725">
         - TableBindings</span></span><br><span data-ttu-id="d21a2-726">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-726">
         - TableCoercion</span></span><br><span data-ttu-id="d21a2-727">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-727">
         - TextBindings</span></span><br><span data-ttu-id="d21a2-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-728">
         - TextCoercion</span></span><br><span data-ttu-id="d21a2-729">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-729">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-730">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="d21a2-730">Office 2016 on Mac</span></span><br><span data-ttu-id="d21a2-731">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-731">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-732">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-732">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d21a2-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="d21a2-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d21a2-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="d21a2-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d21a2-736">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-736">- BindingEvents</span></span><br><span data-ttu-id="d21a2-737">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-737">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-738">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d21a2-738">
         - CustomXmlParts</span></span><br><span data-ttu-id="d21a2-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-739">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-740">
         - File</span></span><br><span data-ttu-id="d21a2-741">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-741">
         - HtmlCoercion</span></span><br><span data-ttu-id="d21a2-742">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-742">
         - MatrixBindings</span></span><br><span data-ttu-id="d21a2-743">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-743">
         - MatrixCoercion</span></span><br><span data-ttu-id="d21a2-744">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-744">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d21a2-745">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-745">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-746">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-746">
         - Selection</span></span><br><span data-ttu-id="d21a2-747">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-747">
         - Settings</span></span><br><span data-ttu-id="d21a2-748">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-748">
         - TableBindings</span></span><br><span data-ttu-id="d21a2-749">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-749">
         - TableCoercion</span></span><br><span data-ttu-id="d21a2-750">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d21a2-750">
         - TextBindings</span></span><br><span data-ttu-id="d21a2-751">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-751">
         - TextCoercion</span></span><br><span data-ttu-id="d21a2-752">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-752">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="d21a2-753">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="d21a2-753">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="d21a2-754">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d21a2-754">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d21a2-755">平台</span><span class="sxs-lookup"><span data-stu-id="d21a2-755">Platform</span></span></th>
    <th><span data-ttu-id="d21a2-756">扩展点</span><span class="sxs-lookup"><span data-stu-id="d21a2-756">Extension points</span></span></th>
    <th><span data-ttu-id="d21a2-757">API 要求集</span><span class="sxs-lookup"><span data-stu-id="d21a2-757">API requirement sets</span></span></th>
    <th><span data-ttu-id="d21a2-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="d21a2-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-759">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="d21a2-759">Office on the web</span></span></td>
    <td> <span data-ttu-id="d21a2-760">- 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-760">- Content</span></span><br><span data-ttu-id="d21a2-761">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-761">
         - TaskPane</span></span><br><span data-ttu-id="d21a2-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d21a2-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d21a2-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d21a2-767">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d21a2-767">- ActiveView</span></span><br><span data-ttu-id="d21a2-768">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-768">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-769">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-769">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-770">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-770">
         - File</span></span><br><span data-ttu-id="d21a2-771">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-771">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-772">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-772">
         - Selection</span></span><br><span data-ttu-id="d21a2-773">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-773">
         - Settings</span></span><br><span data-ttu-id="d21a2-774">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-774">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-775">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="d21a2-775">Office on Windows</span></span><br><span data-ttu-id="d21a2-776">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d21a2-776">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d21a2-777">- 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-777">- Content</span></span><br><span data-ttu-id="d21a2-778">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-778">
         - TaskPane</span></span><br><span data-ttu-id="d21a2-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d21a2-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d21a2-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d21a2-784">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d21a2-784">- ActiveView</span></span><br><span data-ttu-id="d21a2-785">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-785">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-786">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-786">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-787">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-787">
         - File</span></span><br><span data-ttu-id="d21a2-788">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-788">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-789">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-789">
         - Selection</span></span><br><span data-ttu-id="d21a2-790">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-790">
         - Settings</span></span><br><span data-ttu-id="d21a2-791">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-791">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-792">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="d21a2-792">Office 2019 on Windows</span></span><br><span data-ttu-id="d21a2-793">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-793">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-794">- 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-794">- Content</span></span><br><span data-ttu-id="d21a2-795">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-795">
         - TaskPane</span></span><br><span data-ttu-id="d21a2-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d21a2-799">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d21a2-799">- ActiveView</span></span><br><span data-ttu-id="d21a2-800">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-800">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-801">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-801">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-802">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-802">
         - File</span></span><br><span data-ttu-id="d21a2-803">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-803">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-804">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-804">
         - Selection</span></span><br><span data-ttu-id="d21a2-805">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-805">
         - Settings</span></span><br><span data-ttu-id="d21a2-806">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-806">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-807">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="d21a2-807">Office 2016 on Windows</span></span><br><span data-ttu-id="d21a2-808">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-808">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-809">- 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-809">- Content</span></span><br><span data-ttu-id="d21a2-810">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-810">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d21a2-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d21a2-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d21a2-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d21a2-813">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d21a2-813">- ActiveView</span></span><br><span data-ttu-id="d21a2-814">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-814">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-815">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-815">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-816">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-816">
         - File</span></span><br><span data-ttu-id="d21a2-817">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-817">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-818">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-818">
         - Selection</span></span><br><span data-ttu-id="d21a2-819">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-819">
         - Settings</span></span><br><span data-ttu-id="d21a2-820">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-820">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-821">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="d21a2-821">Office 2013 on Windows</span></span><br><span data-ttu-id="d21a2-822">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-822">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-823">- 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-823">- Content</span></span><br><span data-ttu-id="d21a2-824">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-824">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="d21a2-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d21a2-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d21a2-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d21a2-827">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d21a2-827">- ActiveView</span></span><br><span data-ttu-id="d21a2-828">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-828">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-829">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-829">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-830">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-830">
         - File</span></span><br><span data-ttu-id="d21a2-831">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-831">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-832">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-832">
         - Selection</span></span><br><span data-ttu-id="d21a2-833">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-833">
         - Settings</span></span><br><span data-ttu-id="d21a2-834">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-834">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-835">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="d21a2-835">Office on iPad</span></span><br><span data-ttu-id="d21a2-836">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d21a2-836">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d21a2-837">- 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-837">- Content</span></span><br><span data-ttu-id="d21a2-838">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-838">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d21a2-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d21a2-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d21a2-842">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d21a2-842">- ActiveView</span></span><br><span data-ttu-id="d21a2-843">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-843">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-844">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-844">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-845">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-845">
         - File</span></span><br><span data-ttu-id="d21a2-846">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-846">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-847">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-847">
         - Selection</span></span><br><span data-ttu-id="d21a2-848">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-848">
         - Settings</span></span><br><span data-ttu-id="d21a2-849">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-849">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-850">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="d21a2-850">Office on Mac</span></span><br><span data-ttu-id="d21a2-851">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d21a2-851">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="d21a2-852">- 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-852">- Content</span></span><br><span data-ttu-id="d21a2-853">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-853">
         - TaskPane</span></span><br><span data-ttu-id="d21a2-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="d21a2-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="d21a2-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="d21a2-859">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d21a2-859">- ActiveView</span></span><br><span data-ttu-id="d21a2-860">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-860">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-861">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-861">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-862">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-862">
         - File</span></span><br><span data-ttu-id="d21a2-863">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-863">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-864">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-864">
         - Selection</span></span><br><span data-ttu-id="d21a2-865">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-865">
         - Settings</span></span><br><span data-ttu-id="d21a2-866">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-866">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-867">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="d21a2-867">Office 2019 on Mac</span></span><br><span data-ttu-id="d21a2-868">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-868">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-869">- 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-869">- Content</span></span><br><span data-ttu-id="d21a2-870">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-870">
         - TaskPane</span></span><br><span data-ttu-id="d21a2-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d21a2-874">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d21a2-874">- ActiveView</span></span><br><span data-ttu-id="d21a2-875">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-875">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-876">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-876">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-877">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-877">
         - File</span></span><br><span data-ttu-id="d21a2-878">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-878">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-879">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-879">
         - Selection</span></span><br><span data-ttu-id="d21a2-880">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-880">
         - Settings</span></span><br><span data-ttu-id="d21a2-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-881">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-882">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="d21a2-882">Office 2016 on Mac</span></span><br><span data-ttu-id="d21a2-883">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-883">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-884">- 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-884">- Content</span></span><br><span data-ttu-id="d21a2-885">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-885">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d21a2-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d21a2-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="d21a2-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d21a2-888">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d21a2-888">- ActiveView</span></span><br><span data-ttu-id="d21a2-889">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-889">
         - CompressedFile</span></span><br><span data-ttu-id="d21a2-890">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-890">
         - DocumentEvents</span></span><br><span data-ttu-id="d21a2-891">
         - File</span><span class="sxs-lookup"><span data-stu-id="d21a2-891">
         - File</span></span><br><span data-ttu-id="d21a2-892">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d21a2-892">
         - PdfFile</span></span><br><span data-ttu-id="d21a2-893">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-893">
         - Selection</span></span><br><span data-ttu-id="d21a2-894">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-894">
         - Settings</span></span><br><span data-ttu-id="d21a2-895">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-895">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="d21a2-896">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="d21a2-896">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="d21a2-897">OneNote</span><span class="sxs-lookup"><span data-stu-id="d21a2-897">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d21a2-898">平台</span><span class="sxs-lookup"><span data-stu-id="d21a2-898">Platform</span></span></th>
    <th><span data-ttu-id="d21a2-899">扩展点</span><span class="sxs-lookup"><span data-stu-id="d21a2-899">Extension points</span></span></th>
    <th><span data-ttu-id="d21a2-900">API 要求集</span><span class="sxs-lookup"><span data-stu-id="d21a2-900">API requirement sets</span></span></th>
    <th><span data-ttu-id="d21a2-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="d21a2-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-902">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="d21a2-902">Office on the web</span></span></td>
    <td> <span data-ttu-id="d21a2-903">- 内容</span><span class="sxs-lookup"><span data-stu-id="d21a2-903">- Content</span></span><br><span data-ttu-id="d21a2-904">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-904">
         - TaskPane</span></span><br><span data-ttu-id="d21a2-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d21a2-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="d21a2-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="d21a2-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="d21a2-909">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d21a2-909">- DocumentEvents</span></span><br><span data-ttu-id="d21a2-910">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-910">
         - HtmlCoercion</span></span><br><span data-ttu-id="d21a2-911">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d21a2-911">
         - Settings</span></span><br><span data-ttu-id="d21a2-912">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-912">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="d21a2-913">项目</span><span class="sxs-lookup"><span data-stu-id="d21a2-913">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d21a2-914">平台</span><span class="sxs-lookup"><span data-stu-id="d21a2-914">Platform</span></span></th>
    <th><span data-ttu-id="d21a2-915">扩展点</span><span class="sxs-lookup"><span data-stu-id="d21a2-915">Extension points</span></span></th>
    <th><span data-ttu-id="d21a2-916">API 要求集</span><span class="sxs-lookup"><span data-stu-id="d21a2-916">API requirement sets</span></span></th>
    <th><span data-ttu-id="d21a2-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="d21a2-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-918">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="d21a2-918">Office 2019 on Windows</span></span><br><span data-ttu-id="d21a2-919">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-919">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-920">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-920">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d21a2-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d21a2-922">- Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-922">- Selection</span></span><br><span data-ttu-id="d21a2-923">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-923">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-924">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="d21a2-924">Office 2016 on Windows</span></span><br><span data-ttu-id="d21a2-925">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-925">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-926">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-926">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d21a2-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d21a2-928">- Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-928">- Selection</span></span><br><span data-ttu-id="d21a2-929">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-929">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d21a2-930">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="d21a2-930">Office 2013 on Windows</span></span><br><span data-ttu-id="d21a2-931">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d21a2-931">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="d21a2-932">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d21a2-932">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d21a2-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d21a2-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d21a2-934">- Selection</span><span class="sxs-lookup"><span data-stu-id="d21a2-934">- Selection</span></span><br><span data-ttu-id="d21a2-935">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d21a2-935">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="d21a2-936">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d21a2-936">See also</span></span>

- [<span data-ttu-id="d21a2-937">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="d21a2-937">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="d21a2-938">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="d21a2-938">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="d21a2-939">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="d21a2-939">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="d21a2-940">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="d21a2-940">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="d21a2-941">API 参考文档</span><span class="sxs-lookup"><span data-stu-id="d21a2-941">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="d21a2-942">Office 365 ProPlus 的更新历史记录</span><span class="sxs-lookup"><span data-stu-id="d21a2-942">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="d21a2-943">Office 2016 和 2019 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="d21a2-943">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="d21a2-944">Office 2013 更新历史记录（即点即用）</span><span class="sxs-lookup"><span data-stu-id="d21a2-944">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="d21a2-945">Office 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="d21a2-945">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="d21a2-946">Outlook 2010、2013 和 2016 更新历史记录 (MSI)</span><span class="sxs-lookup"><span data-stu-id="d21a2-946">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="d21a2-947">Office for Mac 更新历史记录</span><span class="sxs-lookup"><span data-stu-id="d21a2-947">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="d21a2-948">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="d21a2-948">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
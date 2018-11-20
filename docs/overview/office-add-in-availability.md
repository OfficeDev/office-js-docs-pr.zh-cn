---
title: Office 外接程序主机和平台可用性
description: Excel、Word、Outlook、PowerPoint、OneNote 和项目支持的要求集。
ms.date: 11/07/2018
ms.openlocfilehash: f8d7d9d393531301829b31dd171a5332a0da536b
ms.sourcegitcommit: 9b021af6cb23a58486d6c5c7492be425e309bea1
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/15/2018
ms.locfileid: "26533796"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="10581-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="10581-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="10581-104">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。</span><span class="sxs-lookup"><span data-stu-id="10581-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="10581-105">下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="10581-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="10581-p102">通过 MSI 安装的 Office 2016 的生成号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="10581-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="10581-108">Excel</span><span class="sxs-lookup"><span data-stu-id="10581-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="10581-109">平台</span><span class="sxs-lookup"><span data-stu-id="10581-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="10581-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="10581-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="10581-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="10581-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="10581-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="10581-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="10581-113"> (Office Online)</span></span></td>
    <td> <span data-ttu-id="10581-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-114">- Taskpane</span></span><br><span data-ttu-id="10581-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="10581-115">
        - Content</span></span><br><span data-ttu-id="10581-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="10581-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="10581-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="10581-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="10581-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="10581-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="10581-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="10581-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="10581-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="10581-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="10581-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="10581-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="10581-123">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="10581-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="10581-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="10581-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="10581-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="10581-126">
        -BindingEvents</span></span><br><span data-ttu-id="10581-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-127">
        -CompressedFile</span></span><br><span data-ttu-id="10581-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-128">
        -DocumentEvents</span></span><br><span data-ttu-id="10581-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="10581-129">
        - File</span></span><br><span data-ttu-id="10581-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="10581-130">
        -MatrixBindings</span></span><br><span data-ttu-id="10581-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-131">
        -MatrixCoercion</span></span><br><span data-ttu-id="10581-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-132">
        - Selection</span></span><br><span data-ttu-id="10581-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-133">
        - Settings</span></span><br><span data-ttu-id="10581-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="10581-134">
        -TableBindings</span></span><br><span data-ttu-id="10581-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-135">
        -TableCoercion</span></span><br><span data-ttu-id="10581-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="10581-136">
        -TextBindings</span></span><br><span data-ttu-id="10581-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-137">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-138">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="10581-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="10581-139">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-139">
        - Taskpane</span></span><br><span data-ttu-id="10581-140">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="10581-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="10581-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="10581-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="10581-142">
        -BindingEvents</span></span><br><span data-ttu-id="10581-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-143">
        -CompressedFile</span></span><br><span data-ttu-id="10581-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-144">
        -DocumentEvents</span></span><br><span data-ttu-id="10581-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="10581-145">
        - File</span></span><br><span data-ttu-id="10581-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-146">
        -ImageCoercion</span></span><br><span data-ttu-id="10581-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="10581-147">
        -MatrixBindings</span></span><br><span data-ttu-id="10581-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-148">
        -MatrixCoercion</span></span><br><span data-ttu-id="10581-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-149">
        - Selection</span></span><br><span data-ttu-id="10581-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-150">
        - Settings</span></span><br><span data-ttu-id="10581-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="10581-151">
        -TableBindings</span></span><br><span data-ttu-id="10581-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-152">
        -TableCoercion</span></span><br><span data-ttu-id="10581-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="10581-153">
        -TextBindings</span></span><br><span data-ttu-id="10581-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-154">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-155">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="10581-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="10581-156">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-156">- Taskpane</span></span><br><span data-ttu-id="10581-157">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="10581-157">
        - Content</span></span><br><span data-ttu-id="10581-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="10581-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="10581-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="10581-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="10581-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="10581-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="10581-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="10581-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="10581-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="10581-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="10581-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="10581-165">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="10581-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="10581-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="10581-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="10581-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="10581-168">-BindingEvents</span></span><br><span data-ttu-id="10581-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-169">
        -CompressedFile</span></span><br><span data-ttu-id="10581-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-170">
        -DocumentEvents</span></span><br><span data-ttu-id="10581-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="10581-171">
        - File</span></span><br><span data-ttu-id="10581-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-172">
        -ImageCoercion</span></span><br><span data-ttu-id="10581-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="10581-173">
        -MatrixBindings</span></span><br><span data-ttu-id="10581-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="10581-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-175">
        - Selection</span></span><br><span data-ttu-id="10581-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-176">
        - Settings</span></span><br><span data-ttu-id="10581-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="10581-177">
        -TableBindings</span></span><br><span data-ttu-id="10581-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-178">
        -TableCoercion</span></span><br><span data-ttu-id="10581-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="10581-179">
        -TextBindings</span></span><br><span data-ttu-id="10581-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-181">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="10581-181">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="10581-182">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-182">- Taskpane</span></span><br><span data-ttu-id="10581-183">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="10581-183">
        - Content</span></span><br><span data-ttu-id="10581-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="10581-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="10581-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="10581-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="10581-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="10581-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="10581-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="10581-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="10581-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="10581-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="10581-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="10581-191">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="10581-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="10581-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="10581-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="10581-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="10581-194">-BindingEvents</span></span><br><span data-ttu-id="10581-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-195">
        -CompressedFile</span></span><br><span data-ttu-id="10581-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-196">
        -DocumentEvents</span></span><br><span data-ttu-id="10581-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="10581-197">
        - File</span></span><br><span data-ttu-id="10581-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-198">
        -ImageCoercion</span></span><br><span data-ttu-id="10581-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="10581-199">
        -MatrixBindings</span></span><br><span data-ttu-id="10581-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-200">
        -MatrixCoercion</span></span><br><span data-ttu-id="10581-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-201">
        - Selection</span></span><br><span data-ttu-id="10581-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-202">
        - Settings</span></span><br><span data-ttu-id="10581-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="10581-203">
        -TableBindings</span></span><br><span data-ttu-id="10581-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-204">
        -TableCoercion</span></span><br><span data-ttu-id="10581-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="10581-205">
        -TextBindings</span></span><br><span data-ttu-id="10581-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-206">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-207">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="10581-207">Office for iOS</span></span></td>
    <td><span data-ttu-id="10581-208">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-208">- Taskpane</span></span><br><span data-ttu-id="10581-209">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="10581-209">
        - Content</span></span></td>
    <td><span data-ttu-id="10581-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="10581-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="10581-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="10581-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="10581-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="10581-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="10581-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="10581-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="10581-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="10581-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="10581-216">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="10581-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="10581-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="10581-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="10581-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="10581-219">-BindingEvents</span></span><br><span data-ttu-id="10581-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-220">
        -CompressedFile</span></span><br><span data-ttu-id="10581-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-221">
        -DocumentEvents</span></span><br><span data-ttu-id="10581-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="10581-222">
        - File</span></span><br><span data-ttu-id="10581-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-223">
        -ImageCoercion</span></span><br><span data-ttu-id="10581-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="10581-224">
        -MatrixBindings</span></span><br><span data-ttu-id="10581-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-225">
        -MatrixCoercion</span></span><br><span data-ttu-id="10581-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-226">
        - Selection</span></span><br><span data-ttu-id="10581-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-227">
        - Settings</span></span><br><span data-ttu-id="10581-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="10581-228">
        -TableBindings</span></span><br><span data-ttu-id="10581-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-229">
        -TableCoercion</span></span><br><span data-ttu-id="10581-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="10581-230">
        -TextBindings</span></span><br><span data-ttu-id="10581-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-231">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-232">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="10581-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="10581-233">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-233">- Taskpane</span></span><br><span data-ttu-id="10581-234">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="10581-234">
        - Content</span></span><br><span data-ttu-id="10581-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="10581-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="10581-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="10581-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="10581-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="10581-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="10581-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="10581-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="10581-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="10581-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="10581-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="10581-242">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="10581-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="10581-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="10581-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="10581-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="10581-245">-BindingEvents</span></span><br><span data-ttu-id="10581-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-246">
        -CompressedFile</span></span><br><span data-ttu-id="10581-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-247">
        -DocumentEvents</span></span><br><span data-ttu-id="10581-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="10581-248">
        - File</span></span><br><span data-ttu-id="10581-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-249">
        -ImageCoercion</span></span><br><span data-ttu-id="10581-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="10581-250">
        -MatrixBindings</span></span><br><span data-ttu-id="10581-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-251">
        -MatrixCoercion</span></span><br><span data-ttu-id="10581-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-252">
        -PdfFile</span></span><br><span data-ttu-id="10581-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-253">
        - Selection</span></span><br><span data-ttu-id="10581-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-254">
        - Settings</span></span><br><span data-ttu-id="10581-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="10581-255">
        -TableBindings</span></span><br><span data-ttu-id="10581-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-256">
        -TableCoercion</span></span><br><span data-ttu-id="10581-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="10581-257">
        -TextBindings</span></span><br><span data-ttu-id="10581-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-258">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-259">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="10581-259">Office for Mac</span></span></td>
    <td><span data-ttu-id="10581-260">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-260">- Taskpane</span></span><br><span data-ttu-id="10581-261">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="10581-261">
        - Content</span></span><br><span data-ttu-id="10581-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="10581-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="10581-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="10581-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="10581-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="10581-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="10581-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="10581-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="10581-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="10581-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="10581-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="10581-269">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="10581-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="10581-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="10581-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="10581-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="10581-272">-BindingEvents</span></span><br><span data-ttu-id="10581-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-273">
        -CompressedFile</span></span><br><span data-ttu-id="10581-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-274">
        -DocumentEvents</span></span><br><span data-ttu-id="10581-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="10581-275">
        - File</span></span><br><span data-ttu-id="10581-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-276">
        -ImageCoercion</span></span><br><span data-ttu-id="10581-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="10581-277">
        -MatrixBindings</span></span><br><span data-ttu-id="10581-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-278">
        -MatrixCoercion</span></span><br><span data-ttu-id="10581-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-279">
        -PdfFile</span></span><br><span data-ttu-id="10581-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-280">
        - Selection</span></span><br><span data-ttu-id="10581-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-281">
        - Settings</span></span><br><span data-ttu-id="10581-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="10581-282">
        -TableBindings</span></span><br><span data-ttu-id="10581-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-283">
        -TableCoercion</span></span><br><span data-ttu-id="10581-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="10581-284">
        -TextBindings</span></span><br><span data-ttu-id="10581-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-285">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="10581-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="10581-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="10581-287">平台</span><span class="sxs-lookup"><span data-stu-id="10581-287">Platform</span></span></th>
    <th><span data-ttu-id="10581-288">扩展点</span><span class="sxs-lookup"><span data-stu-id="10581-288">Extension points</span></span></th>
    <th><span data-ttu-id="10581-289">API 要求集</span><span class="sxs-lookup"><span data-stu-id="10581-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="10581-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="10581-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="10581-291"> (Office Online)</span></span></td>
    <td> <span data-ttu-id="10581-292">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="10581-292">- Mail Read</span></span><br><span data-ttu-id="10581-293">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="10581-293">
      - Mail Compose</span></span><br><span data-ttu-id="10581-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="10581-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="10581-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="10581-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="10581-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="10581-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="10581-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="10581-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="10581-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="10581-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="10581-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="10581-302">不可用</span><span class="sxs-lookup"><span data-stu-id="10581-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-303">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="10581-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="10581-304">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="10581-304">- Mail Read</span></span><br><span data-ttu-id="10581-305">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="10581-305">
      - Mail Compose</span></span><br><span data-ttu-id="10581-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="10581-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="10581-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="10581-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="10581-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="10581-311">不可用</span><span class="sxs-lookup"><span data-stu-id="10581-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-312">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="10581-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="10581-313">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="10581-313">- Mail Read</span></span><br><span data-ttu-id="10581-314">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="10581-314">
      - Mail Compose</span></span><br><span data-ttu-id="10581-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="10581-316">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="10581-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="10581-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="10581-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="10581-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="10581-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="10581-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="10581-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="10581-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="10581-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="10581-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="10581-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="10581-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="10581-324">不可用</span><span class="sxs-lookup"><span data-stu-id="10581-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-325">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="10581-325">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="10581-326">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="10581-326">- Mail Read</span></span><br><span data-ttu-id="10581-327">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="10581-327">
      - Mail Compose</span></span><br><span data-ttu-id="10581-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="10581-329">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="10581-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="10581-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="10581-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="10581-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="10581-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="10581-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="10581-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="10581-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="10581-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="10581-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="10581-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="10581-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="10581-337">不可用</span><span class="sxs-lookup"><span data-stu-id="10581-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-338">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="10581-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="10581-339">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="10581-339">- Mail Read</span></span><br><span data-ttu-id="10581-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="10581-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="10581-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="10581-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="10581-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="10581-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="10581-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="10581-346">不可用</span><span class="sxs-lookup"><span data-stu-id="10581-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-347">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="10581-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="10581-348">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="10581-348">- Mail Read</span></span><br><span data-ttu-id="10581-349">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="10581-349">
      - Mail Compose</span></span><br><span data-ttu-id="10581-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="10581-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="10581-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="10581-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="10581-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="10581-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="10581-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="10581-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="10581-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="10581-357">不可用</span><span class="sxs-lookup"><span data-stu-id="10581-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-358">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="10581-358">Office for Mac</span></span></td>
    <td> <span data-ttu-id="10581-359">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="10581-359">- Mail Read</span></span><br><span data-ttu-id="10581-360">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="10581-360">
      - Mail Compose</span></span><br><span data-ttu-id="10581-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="10581-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="10581-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="10581-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="10581-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="10581-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="10581-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="10581-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="10581-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="10581-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="10581-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="10581-369">不可用</span><span class="sxs-lookup"><span data-stu-id="10581-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-370">Office for Android</span><span class="sxs-lookup"><span data-stu-id="10581-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="10581-371">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="10581-371">- Mail Read</span></span><br><span data-ttu-id="10581-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="10581-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="10581-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="10581-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="10581-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="10581-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="10581-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="10581-378">不可用</span><span class="sxs-lookup"><span data-stu-id="10581-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="10581-379">Word</span><span class="sxs-lookup"><span data-stu-id="10581-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="10581-380">平台</span><span class="sxs-lookup"><span data-stu-id="10581-380">Platform</span></span></th>
    <th><span data-ttu-id="10581-381">扩展点</span><span class="sxs-lookup"><span data-stu-id="10581-381">Extension points</span></span></th>
    <th><span data-ttu-id="10581-382">API 要求集</span><span class="sxs-lookup"><span data-stu-id="10581-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="10581-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="10581-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="10581-384"> (Office Online)</span></span></td>
    <td> <span data-ttu-id="10581-385">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-385">- Taskpane</span></span><br><span data-ttu-id="10581-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="10581-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="10581-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="10581-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="10581-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="10581-391">-BindingEvents</span></span><br><span data-ttu-id="10581-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="10581-392">
         -</span></span><br><span data-ttu-id="10581-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-393">
         -DocumentEvents</span></span><br><span data-ttu-id="10581-394">
         - File</span><span class="sxs-lookup"><span data-stu-id="10581-394">
         - File</span></span><br><span data-ttu-id="10581-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-395">
         -HtmlCoercion</span></span><br><span data-ttu-id="10581-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-396">
         -ImageCoercion</span></span><br><span data-ttu-id="10581-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="10581-397">
         -MatrixBindings</span></span><br><span data-ttu-id="10581-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-398">
         -MatrixCoercion</span></span><br><span data-ttu-id="10581-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-399">
         -OoxmlCoercion</span></span><br><span data-ttu-id="10581-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-400">
         -PdfFile</span></span><br><span data-ttu-id="10581-401">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-401">
         - Selection</span></span><br><span data-ttu-id="10581-402">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-402">
         - Settings</span></span><br><span data-ttu-id="10581-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="10581-403">
         -TableBindings</span></span><br><span data-ttu-id="10581-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-404">
         -TableCoercion</span></span><br><span data-ttu-id="10581-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="10581-405">
         -TextBindings</span></span><br><span data-ttu-id="10581-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-406">
         -TextCoercion</span></span><br><span data-ttu-id="10581-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="10581-407">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-408">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="10581-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="10581-409">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-409">- Taskpane</span></span></td>
    <td> <span data-ttu-id="10581-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="10581-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="10581-411">-BindingEvents</span></span><br><span data-ttu-id="10581-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-412">
         -CompressedFile</span></span><br><span data-ttu-id="10581-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="10581-413">
         -</span></span><br><span data-ttu-id="10581-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-414">
         -DocumentEvents</span></span><br><span data-ttu-id="10581-415">
         - File</span><span class="sxs-lookup"><span data-stu-id="10581-415">
         - File</span></span><br><span data-ttu-id="10581-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-416">
         -HtmlCoercion</span></span><br><span data-ttu-id="10581-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-417">
         -ImageCoercion</span></span><br><span data-ttu-id="10581-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="10581-418">
         -MatrixBindings</span></span><br><span data-ttu-id="10581-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-419">
         -MatrixCoercion</span></span><br><span data-ttu-id="10581-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-420">
         -OoxmlCoercion</span></span><br><span data-ttu-id="10581-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-421">
         -PdfFile</span></span><br><span data-ttu-id="10581-422">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-422">
         - Selection</span></span><br><span data-ttu-id="10581-423">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-423">
         - Settings</span></span><br><span data-ttu-id="10581-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="10581-424">
         -TableBindings</span></span><br><span data-ttu-id="10581-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-425">
         -TableCoercion</span></span><br><span data-ttu-id="10581-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="10581-426">
         -TextBindings</span></span><br><span data-ttu-id="10581-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-427">
         -TextCoercion</span></span><br><span data-ttu-id="10581-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="10581-428">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-429">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="10581-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="10581-430">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-430">- Taskpane</span></span><br><span data-ttu-id="10581-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="10581-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="10581-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="10581-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="10581-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="10581-436">-BindingEvents</span></span><br><span data-ttu-id="10581-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-437">
         -CompressedFile</span></span><br><span data-ttu-id="10581-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="10581-438">
         -</span></span><br><span data-ttu-id="10581-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-439">
         -DocumentEvents</span></span><br><span data-ttu-id="10581-440">
         - File</span><span class="sxs-lookup"><span data-stu-id="10581-440">
         - File</span></span><br><span data-ttu-id="10581-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-441">
         -HtmlCoercion</span></span><br><span data-ttu-id="10581-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-442">
         -ImageCoercion</span></span><br><span data-ttu-id="10581-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="10581-443">
         -MatrixBindings</span></span><br><span data-ttu-id="10581-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-444">
         -MatrixCoercion</span></span><br><span data-ttu-id="10581-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-445">
         -OoxmlCoercion</span></span><br><span data-ttu-id="10581-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-446">
         -PdfFile</span></span><br><span data-ttu-id="10581-447">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-447">
         - Selection</span></span><br><span data-ttu-id="10581-448">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-448">
         - Settings</span></span><br><span data-ttu-id="10581-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="10581-449">
         -TableBindings</span></span><br><span data-ttu-id="10581-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-450">
         -TableCoercion</span></span><br><span data-ttu-id="10581-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="10581-451">
         -TextBindings</span></span><br><span data-ttu-id="10581-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-452">
         -TextCoercion</span></span><br><span data-ttu-id="10581-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="10581-453">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-454">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="10581-454">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="10581-455">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-455">- Taskpane</span></span><br><span data-ttu-id="10581-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="10581-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="10581-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="10581-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="10581-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="10581-461">-BindingEvents</span></span><br><span data-ttu-id="10581-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-462">
         -CompressedFile</span></span><br><span data-ttu-id="10581-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="10581-463">
         -</span></span><br><span data-ttu-id="10581-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-464">
         -DocumentEvents</span></span><br><span data-ttu-id="10581-465">
         - File</span><span class="sxs-lookup"><span data-stu-id="10581-465">
         - File</span></span><br><span data-ttu-id="10581-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-466">
         -HtmlCoercion</span></span><br><span data-ttu-id="10581-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-467">
         -ImageCoercion</span></span><br><span data-ttu-id="10581-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="10581-468">
         -MatrixBindings</span></span><br><span data-ttu-id="10581-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-469">
         -MatrixCoercion</span></span><br><span data-ttu-id="10581-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-470">
         -OoxmlCoercion</span></span><br><span data-ttu-id="10581-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-471">
         -PdfFile</span></span><br><span data-ttu-id="10581-472">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-472">
         - Selection</span></span><br><span data-ttu-id="10581-473">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-473">
         - Settings</span></span><br><span data-ttu-id="10581-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="10581-474">
         -TableBindings</span></span><br><span data-ttu-id="10581-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-475">
         -TableCoercion</span></span><br><span data-ttu-id="10581-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="10581-476">
         -TextBindings</span></span><br><span data-ttu-id="10581-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-477">
         -TextCoercion</span></span><br><span data-ttu-id="10581-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="10581-478">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-479">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="10581-479">Office for iOS</span></span></td>
    <td> <span data-ttu-id="10581-480">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-480">- Taskpane</span></span></td>
    <td> <span data-ttu-id="10581-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="10581-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="10581-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="10581-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="10581-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="10581-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="10581-485">-BindingEvents</span></span><br><span data-ttu-id="10581-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-486">
         -CompressedFile</span></span><br><span data-ttu-id="10581-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="10581-487">
         -</span></span><br><span data-ttu-id="10581-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-488">
         -DocumentEvents</span></span><br><span data-ttu-id="10581-489">
         - File</span><span class="sxs-lookup"><span data-stu-id="10581-489">
         - File</span></span><br><span data-ttu-id="10581-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-490">
         -HtmlCoercion</span></span><br><span data-ttu-id="10581-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-491">
         -ImageCoercion</span></span><br><span data-ttu-id="10581-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="10581-492">
         -MatrixBindings</span></span><br><span data-ttu-id="10581-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-493">
         -MatrixCoercion</span></span><br><span data-ttu-id="10581-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-494">
         -OoxmlCoercion</span></span><br><span data-ttu-id="10581-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-495">
         -PdfFile</span></span><br><span data-ttu-id="10581-496">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-496">
         - Selection</span></span><br><span data-ttu-id="10581-497">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-497">
         - Settings</span></span><br><span data-ttu-id="10581-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="10581-498">
         -TableBindings</span></span><br><span data-ttu-id="10581-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-499">
         -TableCoercion</span></span><br><span data-ttu-id="10581-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="10581-500">
         -TextBindings</span></span><br><span data-ttu-id="10581-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-501">
         -TextCoercion</span></span><br><span data-ttu-id="10581-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="10581-502">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-503">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="10581-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="10581-504">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-504">- Taskpane</span></span><br><span data-ttu-id="10581-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="10581-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="10581-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="10581-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="10581-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="10581-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="10581-510">-BindingEvents</span></span><br><span data-ttu-id="10581-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-511">
         -CompressedFile</span></span><br><span data-ttu-id="10581-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="10581-512">
         -</span></span><br><span data-ttu-id="10581-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-513">
         -DocumentEvents</span></span><br><span data-ttu-id="10581-514">
         - File</span><span class="sxs-lookup"><span data-stu-id="10581-514">
         - File</span></span><br><span data-ttu-id="10581-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-515">
         -HtmlCoercion</span></span><br><span data-ttu-id="10581-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-516">
         -ImageCoercion</span></span><br><span data-ttu-id="10581-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="10581-517">
         -MatrixBindings</span></span><br><span data-ttu-id="10581-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-518">
         -MatrixCoercion</span></span><br><span data-ttu-id="10581-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-519">
         -OoxmlCoercion</span></span><br><span data-ttu-id="10581-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-520">
         -PdfFile</span></span><br><span data-ttu-id="10581-521">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-521">
         - Selection</span></span><br><span data-ttu-id="10581-522">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-522">
         - Settings</span></span><br><span data-ttu-id="10581-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="10581-523">
         -TableBindings</span></span><br><span data-ttu-id="10581-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-524">
         -TableCoercion</span></span><br><span data-ttu-id="10581-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="10581-525">
         -TextBindings</span></span><br><span data-ttu-id="10581-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-526">
         -TextCoercion</span></span><br><span data-ttu-id="10581-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="10581-527">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-528">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="10581-528">Office for Mac</span></span></td>
    <td> <span data-ttu-id="10581-529">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-529">- Taskpane</span></span><br><span data-ttu-id="10581-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="10581-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="10581-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="10581-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="10581-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="10581-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="10581-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="10581-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="10581-535">-BindingEvents</span></span><br><span data-ttu-id="10581-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-536">
         -CompressedFile</span></span><br><span data-ttu-id="10581-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="10581-537">
         -</span></span><br><span data-ttu-id="10581-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-538">
         -DocumentEvents</span></span><br><span data-ttu-id="10581-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="10581-539">
         - File</span></span><br><span data-ttu-id="10581-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-540">
         -HtmlCoercion</span></span><br><span data-ttu-id="10581-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-541">
         -ImageCoercion</span></span><br><span data-ttu-id="10581-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="10581-542">
         -MatrixBindings</span></span><br><span data-ttu-id="10581-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-543">
         -MatrixCoercion</span></span><br><span data-ttu-id="10581-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-544">
         -OoxmlCoercion</span></span><br><span data-ttu-id="10581-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-545">
         -PdfFile</span></span><br><span data-ttu-id="10581-546">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-546">
         - Selection</span></span><br><span data-ttu-id="10581-547">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-547">
         - Settings</span></span><br><span data-ttu-id="10581-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="10581-548">
         -TableBindings</span></span><br><span data-ttu-id="10581-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-549">
         -TableCoercion</span></span><br><span data-ttu-id="10581-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="10581-550">
         -TextBindings</span></span><br><span data-ttu-id="10581-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-551">
         -TextCoercion</span></span><br><span data-ttu-id="10581-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="10581-552">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="10581-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="10581-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="10581-554">平台</span><span class="sxs-lookup"><span data-stu-id="10581-554">Platform</span></span></th>
    <th><span data-ttu-id="10581-555">扩展点</span><span class="sxs-lookup"><span data-stu-id="10581-555">Extension points</span></span></th>
    <th><span data-ttu-id="10581-556">API 要求集</span><span class="sxs-lookup"><span data-stu-id="10581-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="10581-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="10581-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="10581-558"> (Office Online)</span></span></td>
    <td> <span data-ttu-id="10581-559">- 内容</span><span class="sxs-lookup"><span data-stu-id="10581-559">- Content</span></span><br><span data-ttu-id="10581-560">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-560">
         - Taskpane</span></span><br><span data-ttu-id="10581-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="10581-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="10581-563">-</span></span><br><span data-ttu-id="10581-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-564">
         -CompressedFile</span></span><br><span data-ttu-id="10581-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-565">
         -DocumentEvents</span></span><br><span data-ttu-id="10581-566">
         - File</span><span class="sxs-lookup"><span data-stu-id="10581-566">
         - File</span></span><br><span data-ttu-id="10581-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-567">
         -ImageCoercion</span></span><br><span data-ttu-id="10581-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-568">
         -PdfFile</span></span><br><span data-ttu-id="10581-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-569">
         - Selection</span></span><br><span data-ttu-id="10581-570">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-570">
         - Settings</span></span><br><span data-ttu-id="10581-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-571">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-572">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="10581-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="10581-573">- 内容</span><span class="sxs-lookup"><span data-stu-id="10581-573">- Content</span></span><br><span data-ttu-id="10581-574">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-574">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="10581-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="10581-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="10581-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="10581-576">-</span></span><br><span data-ttu-id="10581-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-577">
         -CompressedFile</span></span><br><span data-ttu-id="10581-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-578">
         -DocumentEvents</span></span><br><span data-ttu-id="10581-579">
         - File</span><span class="sxs-lookup"><span data-stu-id="10581-579">
         - File</span></span><br><span data-ttu-id="10581-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-580">
         -ImageCoercion</span></span><br><span data-ttu-id="10581-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-581">
         -PdfFile</span></span><br><span data-ttu-id="10581-582">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-582">
         - Selection</span></span><br><span data-ttu-id="10581-583">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-583">
         - Settings</span></span><br><span data-ttu-id="10581-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-584">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-585">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="10581-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="10581-586">- 内容</span><span class="sxs-lookup"><span data-stu-id="10581-586">- Content</span></span><br><span data-ttu-id="10581-587">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-587">
         - Taskpane</span></span><br><span data-ttu-id="10581-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="10581-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="10581-590">-</span></span><br><span data-ttu-id="10581-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-591">
         -CompressedFile</span></span><br><span data-ttu-id="10581-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-592">
         -DocumentEvents</span></span><br><span data-ttu-id="10581-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="10581-593">
         - File</span></span><br><span data-ttu-id="10581-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-594">
         -ImageCoercion</span></span><br><span data-ttu-id="10581-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-595">
         -PdfFile</span></span><br><span data-ttu-id="10581-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-596">
         - Selection</span></span><br><span data-ttu-id="10581-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-597">
         - Settings</span></span><br><span data-ttu-id="10581-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-598">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-599">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="10581-599">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="10581-600">- 内容</span><span class="sxs-lookup"><span data-stu-id="10581-600">- Content</span></span><br><span data-ttu-id="10581-601">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-601">
         - Taskpane</span></span><br><span data-ttu-id="10581-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="10581-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="10581-604">-</span></span><br><span data-ttu-id="10581-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-605">
         -CompressedFile</span></span><br><span data-ttu-id="10581-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-606">
         -DocumentEvents</span></span><br><span data-ttu-id="10581-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="10581-607">
         - File</span></span><br><span data-ttu-id="10581-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-608">
         -ImageCoercion</span></span><br><span data-ttu-id="10581-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-609">
         -PdfFile</span></span><br><span data-ttu-id="10581-610">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-610">
         - Selection</span></span><br><span data-ttu-id="10581-611">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-611">
         - Settings</span></span><br><span data-ttu-id="10581-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-612">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-613">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="10581-613">Office for iOS</span></span></td>
    <td> <span data-ttu-id="10581-614">- 内容</span><span class="sxs-lookup"><span data-stu-id="10581-614">- Content</span></span><br><span data-ttu-id="10581-615">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-615">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="10581-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="10581-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="10581-617">-</span></span><br><span data-ttu-id="10581-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-618">
         -CompressedFile</span></span><br><span data-ttu-id="10581-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-619">
         -DocumentEvents</span></span><br><span data-ttu-id="10581-620">
         - File</span><span class="sxs-lookup"><span data-stu-id="10581-620">
         - File</span></span><br><span data-ttu-id="10581-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-621">
         -PdfFile</span></span><br><span data-ttu-id="10581-622">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-622">
         - Selection</span></span><br><span data-ttu-id="10581-623">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-623">
         - Settings</span></span><br><span data-ttu-id="10581-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-624">
         -TextCoercion</span></span><br><span data-ttu-id="10581-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-625">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-626">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="10581-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="10581-627">- 内容</span><span class="sxs-lookup"><span data-stu-id="10581-627">- Content</span></span><br><span data-ttu-id="10581-628">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-628">
         - Taskpane</span></span><br><span data-ttu-id="10581-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="10581-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="10581-631">-</span></span><br><span data-ttu-id="10581-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-632">
         -CompressedFile</span></span><br><span data-ttu-id="10581-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-633">
         -DocumentEvents</span></span><br><span data-ttu-id="10581-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="10581-634">
         - File</span></span><br><span data-ttu-id="10581-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-635">
         -ImageCoercion</span></span><br><span data-ttu-id="10581-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-636">
         -PdfFile</span></span><br><span data-ttu-id="10581-637">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-637">
         - Selection</span></span><br><span data-ttu-id="10581-638">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-638">
         - Settings</span></span><br><span data-ttu-id="10581-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-639">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-640">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="10581-640">Office for Mac</span></span></td>
    <td> <span data-ttu-id="10581-641">- 内容</span><span class="sxs-lookup"><span data-stu-id="10581-641">- Content</span></span><br><span data-ttu-id="10581-642">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-642">
         - Taskpane</span></span><br><span data-ttu-id="10581-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="10581-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="10581-645">-</span></span><br><span data-ttu-id="10581-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="10581-646">
         -CompressedFile</span></span><br><span data-ttu-id="10581-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-647">
         -DocumentEvents</span></span><br><span data-ttu-id="10581-648">
         - File</span><span class="sxs-lookup"><span data-stu-id="10581-648">
         - File</span></span><br><span data-ttu-id="10581-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-649">
         -ImageCoercion</span></span><br><span data-ttu-id="10581-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="10581-650">
         -PdfFile</span></span><br><span data-ttu-id="10581-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="10581-651">
         - Selection</span></span><br><span data-ttu-id="10581-652">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-652">
         - Settings</span></span><br><span data-ttu-id="10581-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-653">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="10581-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="10581-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="10581-655">平台</span><span class="sxs-lookup"><span data-stu-id="10581-655">Platform</span></span></th>
    <th><span data-ttu-id="10581-656">扩展点</span><span class="sxs-lookup"><span data-stu-id="10581-656">Extension points</span></span></th>
    <th><span data-ttu-id="10581-657">API 要求集</span><span class="sxs-lookup"><span data-stu-id="10581-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="10581-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="10581-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="10581-659"> (Office Online)</span></span></td>
    <td> <span data-ttu-id="10581-660">- 内容</span><span class="sxs-lookup"><span data-stu-id="10581-660">- Content</span></span><br><span data-ttu-id="10581-661">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-661">
         - Taskpane</span></span><br><span data-ttu-id="10581-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="10581-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="10581-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="10581-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="10581-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="10581-665">-DocumentEvents</span></span><br><span data-ttu-id="10581-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-666">
         -HtmlCoercion</span></span><br><span data-ttu-id="10581-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-667">
         -ImageCoercion</span></span><br><span data-ttu-id="10581-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="10581-668">
         - Settings</span></span><br><span data-ttu-id="10581-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-669">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="10581-670">项目</span><span class="sxs-lookup"><span data-stu-id="10581-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="10581-671">平台</span><span class="sxs-lookup"><span data-stu-id="10581-671">Platform</span></span></th>
    <th><span data-ttu-id="10581-672">扩展点</span><span class="sxs-lookup"><span data-stu-id="10581-672">Extension points</span></span></th>
    <th><span data-ttu-id="10581-673">API 要求集</span><span class="sxs-lookup"><span data-stu-id="10581-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="10581-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="10581-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-675">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="10581-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="10581-676">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-676">- Taskpane</span></span></td>
    <td> <span data-ttu-id="10581-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="10581-678">- Selection</span><span class="sxs-lookup"><span data-stu-id="10581-678">- Selection</span></span><br><span data-ttu-id="10581-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-679">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-680">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="10581-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="10581-681">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-681">- Taskpane</span></span></td>
    <td> <span data-ttu-id="10581-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="10581-683">- Selection</span><span class="sxs-lookup"><span data-stu-id="10581-683">- Selection</span></span><br><span data-ttu-id="10581-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-684">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="10581-685">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="10581-685">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="10581-686">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="10581-686">- Taskpane</span></span></td>
    <td> <span data-ttu-id="10581-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="10581-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="10581-688">- Selection</span><span class="sxs-lookup"><span data-stu-id="10581-688">- Selection</span></span><br><span data-ttu-id="10581-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="10581-689">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="10581-690">另请参阅</span><span class="sxs-lookup"><span data-stu-id="10581-690">See also</span></span>

- [<span data-ttu-id="10581-691">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="10581-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="10581-692">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="10581-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="10581-693">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="10581-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="10581-694">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="10581-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)

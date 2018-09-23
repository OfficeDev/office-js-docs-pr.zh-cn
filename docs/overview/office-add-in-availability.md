---
title: Office 加载项主机和平台可用性
description: Excel、Word、Outlook、PowerPoint 和 OneNote 支持的要求集。
ms.date: 09/19/2018
ms.openlocfilehash: 09fb72c88bd0496c413f94b7ba4149192380d664
ms.sourcegitcommit: e7e4d08569a01c69168bb005188e9a1e628304b9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/22/2018
ms.locfileid: "24967702"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="9df16-103">Office 加载项主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="9df16-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="9df16-104">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。</span><span class="sxs-lookup"><span data-stu-id="9df16-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="9df16-105">下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="9df16-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="9df16-106">如果表格单元格内有星号 (\*)，表示我们正在完善它。</span><span class="sxs-lookup"><span data-stu-id="9df16-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="9df16-107">有关 Project 或 Access 要求集，请参阅 [Office 通用要求集](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="9df16-107">For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="9df16-p103">通过 MSI 安装的 Office 2016 的生成号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="9df16-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="9df16-110">Excel</span><span class="sxs-lookup"><span data-stu-id="9df16-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="9df16-111">平台</span><span class="sxs-lookup"><span data-stu-id="9df16-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="9df16-112">扩展点</span><span class="sxs-lookup"><span data-stu-id="9df16-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="9df16-113">API 要求集</span><span class="sxs-lookup"><span data-stu-id="9df16-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="9df16-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="9df16-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="9df16-115">Office Online</span></span></td>
    <td> <span data-ttu-id="9df16-116">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-116">- Taskpane</span></span><br><span data-ttu-id="9df16-117">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="9df16-117">
        - Content</span></span><br><span data-ttu-id="9df16-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="9df16-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="9df16-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9df16-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9df16-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9df16-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9df16-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9df16-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9df16-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9df16-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9df16-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9df16-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9df16-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9df16-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9df16-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="9df16-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9df16-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-127">
        -BindingEvents</span></span><br><span data-ttu-id="9df16-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9df16-128">
        -CompressedFile</span></span><br><span data-ttu-id="9df16-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-129">
        -DocumentEvents</span></span><br><span data-ttu-id="9df16-130">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="9df16-130">
        - File</span></span><br><span data-ttu-id="9df16-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-131">
        -MatrixBindings</span></span><br><span data-ttu-id="9df16-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-132">
        -MatrixCoercion</span></span><br><span data-ttu-id="9df16-133">
        - 选择</span><span class="sxs-lookup"><span data-stu-id="9df16-133">
        - Selection</span></span><br><span data-ttu-id="9df16-134">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-134">
        - Settings</span></span><br><span data-ttu-id="9df16-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-135">
        -TableBindings</span></span><br><span data-ttu-id="9df16-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-136">
        -TableCoercion</span></span><br><span data-ttu-id="9df16-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-137">
        -TextBindings</span></span><br><span data-ttu-id="9df16-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-138">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-139">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="9df16-139">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="9df16-140">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-140">
        - Taskpane</span></span><br><span data-ttu-id="9df16-141">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="9df16-141">
        - Content</span></span></td>
    <td>  <span data-ttu-id="9df16-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9df16-143">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-143">
        -BindingEvents</span></span><br><span data-ttu-id="9df16-144">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9df16-144">
        -CompressedFile</span></span><br><span data-ttu-id="9df16-145">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-145">
        -DocumentEvents</span></span><br><span data-ttu-id="9df16-146">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="9df16-146">
        - File</span></span><br><span data-ttu-id="9df16-147">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-147">
        -ImageCoercion</span></span><br><span data-ttu-id="9df16-148">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-148">
        -MatrixBindings</span></span><br><span data-ttu-id="9df16-149">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-149">
        -MatrixCoercion</span></span><br><span data-ttu-id="9df16-150">
        - 选择</span><span class="sxs-lookup"><span data-stu-id="9df16-150">
        - Selection</span></span><br><span data-ttu-id="9df16-151">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-151">
        - Settings</span></span><br><span data-ttu-id="9df16-152">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-152">
        -TableBindings</span></span><br><span data-ttu-id="9df16-153">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-153">
        -TableCoercion</span></span><br><span data-ttu-id="9df16-154">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-154">
        -TextBindings</span></span><br><span data-ttu-id="9df16-155">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-155">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-156">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="9df16-156">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="9df16-157">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-157">- Taskpane</span></span><br><span data-ttu-id="9df16-158">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="9df16-158">
        - Content</span></span><br><span data-ttu-id="9df16-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9df16-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9df16-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9df16-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9df16-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9df16-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9df16-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9df16-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9df16-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9df16-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9df16-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9df16-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9df16-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9df16-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9df16-166">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="9df16-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9df16-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-168">-BindingEvents</span></span><br><span data-ttu-id="9df16-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9df16-169">
        -CompressedFile</span></span><br><span data-ttu-id="9df16-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-170">
        -DocumentEvents</span></span><br><span data-ttu-id="9df16-171">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="9df16-171">
        - File</span></span><br><span data-ttu-id="9df16-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-172">
        -ImageCoercion</span></span><br><span data-ttu-id="9df16-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-173">
        -MatrixBindings</span></span><br><span data-ttu-id="9df16-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="9df16-175">
        - 选择</span><span class="sxs-lookup"><span data-stu-id="9df16-175">
        - Selection</span></span><br><span data-ttu-id="9df16-176">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-176">
        - Settings</span></span><br><span data-ttu-id="9df16-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-177">
        -TableBindings</span></span><br><span data-ttu-id="9df16-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-178">
        -TableCoercion</span></span><br><span data-ttu-id="9df16-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-179">
        -TextBindings</span></span><br><span data-ttu-id="9df16-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-181">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="9df16-181">Office for iOS</span></span></td>
    <td><span data-ttu-id="9df16-182">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-182">- Taskpane</span></span><br><span data-ttu-id="9df16-183">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="9df16-183">
        - Content</span></span></td>
    <td><span data-ttu-id="9df16-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9df16-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9df16-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9df16-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9df16-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9df16-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9df16-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9df16-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9df16-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9df16-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9df16-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9df16-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9df16-190">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="9df16-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9df16-192">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-192">-BindingEvents</span></span><br><span data-ttu-id="9df16-193">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9df16-193">
        -CompressedFile</span></span><br><span data-ttu-id="9df16-194">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-194">
        -DocumentEvents</span></span><br><span data-ttu-id="9df16-195">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="9df16-195">
        - File</span></span><br><span data-ttu-id="9df16-196">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-196">
        -ImageCoercion</span></span><br><span data-ttu-id="9df16-197">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-197">
        -MatrixBindings</span></span><br><span data-ttu-id="9df16-198">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-198">
        -MatrixCoercion</span></span><br><span data-ttu-id="9df16-199">
        - 选择</span><span class="sxs-lookup"><span data-stu-id="9df16-199">
        - Selection</span></span><br><span data-ttu-id="9df16-200">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-200">
        - Settings</span></span><br><span data-ttu-id="9df16-201">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-201">
        -TableBindings</span></span><br><span data-ttu-id="9df16-202">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-202">
        -TableCoercion</span></span><br><span data-ttu-id="9df16-203">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-203">
        -TextBindings</span></span><br><span data-ttu-id="9df16-204">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-204">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-205">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="9df16-205">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="9df16-206">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-206">- Taskpane</span></span><br><span data-ttu-id="9df16-207">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="9df16-207">
        - Content</span></span><br><span data-ttu-id="9df16-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9df16-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9df16-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9df16-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9df16-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9df16-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9df16-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9df16-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9df16-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9df16-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9df16-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9df16-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9df16-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9df16-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9df16-215">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="9df16-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9df16-217">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-217">-BindingEvents</span></span><br><span data-ttu-id="9df16-218">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9df16-218">
        -CompressedFile</span></span><br><span data-ttu-id="9df16-219">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-219">
        -DocumentEvents</span></span><br><span data-ttu-id="9df16-220">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="9df16-220">
        - File</span></span><br><span data-ttu-id="9df16-221">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-221">
        -ImageCoercion</span></span><br><span data-ttu-id="9df16-222">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-222">
        -MatrixBindings</span></span><br><span data-ttu-id="9df16-223">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-223">
        -MatrixCoercion</span></span><br><span data-ttu-id="9df16-224">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9df16-224">
        -PdfFile</span></span><br><span data-ttu-id="9df16-225">
        - 选择</span><span class="sxs-lookup"><span data-stu-id="9df16-225">
        - Selection</span></span><br><span data-ttu-id="9df16-226">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-226">
        - Settings</span></span><br><span data-ttu-id="9df16-227">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-227">
        -TableBindings</span></span><br><span data-ttu-id="9df16-228">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-228">
        -TableCoercion</span></span><br><span data-ttu-id="9df16-229">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-229">
        -TextBindings</span></span><br><span data-ttu-id="9df16-230">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-230">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="9df16-231">Outlook</span><span class="sxs-lookup"><span data-stu-id="9df16-231">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9df16-232">平台</span><span class="sxs-lookup"><span data-stu-id="9df16-232">Platform</span></span></th>
    <th><span data-ttu-id="9df16-233">扩展点</span><span class="sxs-lookup"><span data-stu-id="9df16-233">Extension points</span></span></th>
    <th><span data-ttu-id="9df16-234">API 要求集</span><span class="sxs-lookup"><span data-stu-id="9df16-234">API requirement sets</span></span></th>
    <th><span data-ttu-id="9df16-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="9df16-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-236">Office Online</span><span class="sxs-lookup"><span data-stu-id="9df16-236">Office Online</span></span></td>
    <td> <span data-ttu-id="9df16-237">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="9df16-237">- Mail Read</span></span><br><span data-ttu-id="9df16-238">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="9df16-238">
      - Mail Compose</span></span><br><span data-ttu-id="9df16-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9df16-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9df16-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9df16-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9df16-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9df16-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9df16-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9df16-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9df16-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9df16-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9df16-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9df16-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9df16-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9df16-246">不适用</span><span class="sxs-lookup"><span data-stu-id="9df16-246">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-247">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="9df16-247">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9df16-248">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="9df16-248">- Mail Read</span></span><br><span data-ttu-id="9df16-249">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="9df16-249">
      - Mail Compose</span></span><br><span data-ttu-id="9df16-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9df16-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9df16-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9df16-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9df16-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9df16-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9df16-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9df16-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9df16-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="9df16-255">不适用</span><span class="sxs-lookup"><span data-stu-id="9df16-255">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-256">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="9df16-256">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9df16-257">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="9df16-257">- Mail Read</span></span><br><span data-ttu-id="9df16-258">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="9df16-258">
      - Mail Compose</span></span><br><span data-ttu-id="9df16-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9df16-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="9df16-260">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="9df16-260">
      - Modules</span></span></td>
    <td> <span data-ttu-id="9df16-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9df16-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9df16-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9df16-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9df16-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9df16-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9df16-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9df16-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9df16-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9df16-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9df16-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9df16-267">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9df16-267">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="9df16-268">不适用</span><span class="sxs-lookup"><span data-stu-id="9df16-268">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-269">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="9df16-269">Office for iOS</span></span></td>
    <td> <span data-ttu-id="9df16-270">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="9df16-270">- Mail Read</span></span><br><span data-ttu-id="9df16-271">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9df16-271">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9df16-272">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-272">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9df16-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9df16-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9df16-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9df16-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9df16-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9df16-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9df16-276">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9df16-276">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="9df16-277">不适用</span><span class="sxs-lookup"><span data-stu-id="9df16-277">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-278">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="9df16-278">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="9df16-279">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="9df16-279">- Mail Read</span></span><br><span data-ttu-id="9df16-280">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="9df16-280">
      - Mail Compose</span></span><br><span data-ttu-id="9df16-281">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9df16-281">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9df16-282">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-282">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9df16-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9df16-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9df16-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9df16-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9df16-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9df16-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9df16-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9df16-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9df16-287">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9df16-287">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9df16-288">不适用</span><span class="sxs-lookup"><span data-stu-id="9df16-288">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-289">Office for Android</span><span class="sxs-lookup"><span data-stu-id="9df16-289">Office for Android</span></span></td>
    <td> <span data-ttu-id="9df16-290">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="9df16-290">- Mail Read</span></span><br><span data-ttu-id="9df16-291">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9df16-291">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9df16-292">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-292">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9df16-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9df16-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9df16-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9df16-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9df16-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9df16-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9df16-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9df16-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="9df16-297">不适用</span><span class="sxs-lookup"><span data-stu-id="9df16-297">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="9df16-298">Word</span><span class="sxs-lookup"><span data-stu-id="9df16-298">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9df16-299">平台</span><span class="sxs-lookup"><span data-stu-id="9df16-299">Platform</span></span></th>
    <th><span data-ttu-id="9df16-300">扩展点</span><span class="sxs-lookup"><span data-stu-id="9df16-300">Extension points</span></span></th>
    <th><span data-ttu-id="9df16-301">API 要求集</span><span class="sxs-lookup"><span data-stu-id="9df16-301">API requirement sets</span></span></th>
    <th><span data-ttu-id="9df16-302"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="9df16-302"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-303">Office Online</span><span class="sxs-lookup"><span data-stu-id="9df16-303">Office Online</span></span></td>
    <td> <span data-ttu-id="9df16-304">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-304">- Taskpane</span></span><br><span data-ttu-id="9df16-305">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9df16-305">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9df16-306">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-306">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9df16-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9df16-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9df16-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9df16-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9df16-309">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-309">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9df16-310">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-310">-BindingEvents</span></span><br><span data-ttu-id="9df16-311">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9df16-311">
         -</span></span><br><span data-ttu-id="9df16-312">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-312">
         -DocumentEvents</span></span><br><span data-ttu-id="9df16-313">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9df16-313">
         - File</span></span><br><span data-ttu-id="9df16-314">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-314">
         -HtmlCoercion</span></span><br><span data-ttu-id="9df16-315">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-315">
         -ImageCoercion</span></span><br><span data-ttu-id="9df16-316">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-316">
         -MatrixBindings</span></span><br><span data-ttu-id="9df16-317">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-317">
         -MatrixCoercion</span></span><br><span data-ttu-id="9df16-318">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-318">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9df16-319">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9df16-319">
         -PdfFile</span></span><br><span data-ttu-id="9df16-320">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9df16-320">
         - Selection</span></span><br><span data-ttu-id="9df16-321">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-321">
         - Settings</span></span><br><span data-ttu-id="9df16-322">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-322">
         -TableBindings</span></span><br><span data-ttu-id="9df16-323">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-323">
         -TableCoercion</span></span><br><span data-ttu-id="9df16-324">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-324">
         -TextBindings</span></span><br><span data-ttu-id="9df16-325">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-325">
         -TextCoercion</span></span><br><span data-ttu-id="9df16-326">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9df16-326">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-327">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="9df16-327">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9df16-328">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-328">- Taskpane</span></span></td>
    <td> <span data-ttu-id="9df16-329">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-329">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9df16-330">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-330">-BindingEvents</span></span><br><span data-ttu-id="9df16-331">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9df16-331">
         -CompressedFile</span></span><br><span data-ttu-id="9df16-332">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9df16-332">
         -</span></span><br><span data-ttu-id="9df16-333">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-333">
         -DocumentEvents</span></span><br><span data-ttu-id="9df16-334">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9df16-334">
         - File</span></span><br><span data-ttu-id="9df16-335">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-335">
         -HtmlCoercion</span></span><br><span data-ttu-id="9df16-336">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-336">
         -ImageCoercion</span></span><br><span data-ttu-id="9df16-337">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-337">
         -MatrixBindings</span></span><br><span data-ttu-id="9df16-338">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-338">
         -MatrixCoercion</span></span><br><span data-ttu-id="9df16-339">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-339">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9df16-340">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9df16-340">
         -PdfFile</span></span><br><span data-ttu-id="9df16-341">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9df16-341">
         - Selection</span></span><br><span data-ttu-id="9df16-342">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-342">
         - Settings</span></span><br><span data-ttu-id="9df16-343">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-343">
         -TableBindings</span></span><br><span data-ttu-id="9df16-344">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-344">
         -TableCoercion</span></span><br><span data-ttu-id="9df16-345">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-345">
         -TextBindings</span></span><br><span data-ttu-id="9df16-346">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-346">
         -TextCoercion</span></span><br><span data-ttu-id="9df16-347">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9df16-347">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-348">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="9df16-348">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9df16-349">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-349">- Taskpane</span></span><br><span data-ttu-id="9df16-350">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9df16-350">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9df16-351">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-351">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9df16-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9df16-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9df16-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9df16-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9df16-354">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-354">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9df16-355">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-355">-BindingEvents</span></span><br><span data-ttu-id="9df16-356">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9df16-356">
         -CompressedFile</span></span><br><span data-ttu-id="9df16-357">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9df16-357">
         -</span></span><br><span data-ttu-id="9df16-358">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-358">
         -DocumentEvents</span></span><br><span data-ttu-id="9df16-359">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9df16-359">
         - File</span></span><br><span data-ttu-id="9df16-360">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-360">
         -HtmlCoercion</span></span><br><span data-ttu-id="9df16-361">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-361">
         -ImageCoercion</span></span><br><span data-ttu-id="9df16-362">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-362">
         -MatrixBindings</span></span><br><span data-ttu-id="9df16-363">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-363">
         -MatrixCoercion</span></span><br><span data-ttu-id="9df16-364">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-364">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9df16-365">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9df16-365">
         -PdfFile</span></span><br><span data-ttu-id="9df16-366">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9df16-366">
         - Selection</span></span><br><span data-ttu-id="9df16-367">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-367">
         - Settings</span></span><br><span data-ttu-id="9df16-368">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-368">
         -TableBindings</span></span><br><span data-ttu-id="9df16-369">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-369">
         -TableCoercion</span></span><br><span data-ttu-id="9df16-370">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-370">
         -TextBindings</span></span><br><span data-ttu-id="9df16-371">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-371">
         -TextCoercion</span></span><br><span data-ttu-id="9df16-372">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9df16-372">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-373">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="9df16-373">Office for iOS</span></span></td>
    <td> <span data-ttu-id="9df16-374">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-374">- Taskpane</span></span></td>
    <td> <span data-ttu-id="9df16-375">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-375">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9df16-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9df16-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9df16-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9df16-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9df16-378">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="9df16-378">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="9df16-379">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-379">-BindingEvents</span></span><br><span data-ttu-id="9df16-380">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9df16-380">
         -CompressedFile</span></span><br><span data-ttu-id="9df16-381">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9df16-381">
         -</span></span><br><span data-ttu-id="9df16-382">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-382">
         -DocumentEvents</span></span><br><span data-ttu-id="9df16-383">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9df16-383">
         - File</span></span><br><span data-ttu-id="9df16-384">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-384">
         -HtmlCoercion</span></span><br><span data-ttu-id="9df16-385">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-385">
         -ImageCoercion</span></span><br><span data-ttu-id="9df16-386">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-386">
         -MatrixBindings</span></span><br><span data-ttu-id="9df16-387">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-387">
         -MatrixCoercion</span></span><br><span data-ttu-id="9df16-388">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-388">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9df16-389">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9df16-389">
         -PdfFile</span></span><br><span data-ttu-id="9df16-390">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9df16-390">
         - Selection</span></span><br><span data-ttu-id="9df16-391">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-391">
         - Settings</span></span><br><span data-ttu-id="9df16-392">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-392">
         -TableBindings</span></span><br><span data-ttu-id="9df16-393">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-393">
         -TableCoercion</span></span><br><span data-ttu-id="9df16-394">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-394">
         -TextBindings</span></span><br><span data-ttu-id="9df16-395">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-395">
         -TextCoercion</span></span><br><span data-ttu-id="9df16-396">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9df16-396">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-397">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="9df16-397">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="9df16-398">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-398">- Taskpane</span></span><br><span data-ttu-id="9df16-399">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9df16-399">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9df16-400">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-400">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9df16-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9df16-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9df16-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9df16-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9df16-403">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="9df16-403">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="9df16-404">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-404">-BindingEvents</span></span><br><span data-ttu-id="9df16-405">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9df16-405">
         -CompressedFile</span></span><br><span data-ttu-id="9df16-406">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9df16-406">
         -</span></span><br><span data-ttu-id="9df16-407">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-407">
         -DocumentEvents</span></span><br><span data-ttu-id="9df16-408">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9df16-408">
         - File</span></span><br><span data-ttu-id="9df16-409">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-409">
         -HtmlCoercion</span></span><br><span data-ttu-id="9df16-410">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-410">
         -ImageCoercion</span></span><br><span data-ttu-id="9df16-411">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-411">
         -MatrixBindings</span></span><br><span data-ttu-id="9df16-412">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-412">
         -MatrixCoercion</span></span><br><span data-ttu-id="9df16-413">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-413">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9df16-414">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9df16-414">
         -PdfFile</span></span><br><span data-ttu-id="9df16-415">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9df16-415">
         - Selection</span></span><br><span data-ttu-id="9df16-416">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-416">
         - Settings</span></span><br><span data-ttu-id="9df16-417">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-417">
         -TableBindings</span></span><br><span data-ttu-id="9df16-418">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-418">
         -TableCoercion</span></span><br><span data-ttu-id="9df16-419">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9df16-419">
         -TextBindings</span></span><br><span data-ttu-id="9df16-420">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-420">
         -TextCoercion</span></span><br><span data-ttu-id="9df16-421">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9df16-421">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="9df16-422">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="9df16-422">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9df16-423">平台</span><span class="sxs-lookup"><span data-stu-id="9df16-423">Platform</span></span></th>
    <th><span data-ttu-id="9df16-424">扩展点</span><span class="sxs-lookup"><span data-stu-id="9df16-424">Extension points</span></span></th>
    <th><span data-ttu-id="9df16-425">API 要求集</span><span class="sxs-lookup"><span data-stu-id="9df16-425">API requirement sets</span></span></th>
    <th><span data-ttu-id="9df16-426"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="9df16-426"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-427">Office Online</span><span class="sxs-lookup"><span data-stu-id="9df16-427">Office Online</span></span></td>
    <td> <span data-ttu-id="9df16-428">- 内容</span><span class="sxs-lookup"><span data-stu-id="9df16-428">- Content</span></span><br><span data-ttu-id="9df16-429">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-429">
         - Taskpane</span></span><br><span data-ttu-id="9df16-430">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9df16-430">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9df16-431">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-431">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9df16-432">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9df16-432">-ActiveView</span></span><br><span data-ttu-id="9df16-433">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9df16-433">
         -CompressedFile</span></span><br><span data-ttu-id="9df16-434">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-434">
         -DocumentEvents</span></span><br><span data-ttu-id="9df16-435">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9df16-435">
         - File</span></span><br><span data-ttu-id="9df16-436">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-436">
         -ImageCoercion</span></span><br><span data-ttu-id="9df16-437">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9df16-437">
         -PdfFile</span></span><br><span data-ttu-id="9df16-438">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9df16-438">
         - Selection</span></span><br><span data-ttu-id="9df16-439">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-439">
         - Settings</span></span><br><span data-ttu-id="9df16-440">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-440">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-441">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="9df16-441">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9df16-442">- 内容</span><span class="sxs-lookup"><span data-stu-id="9df16-442">- Content</span></span><br><span data-ttu-id="9df16-443">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-443">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="9df16-444">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="9df16-444">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="9df16-445">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9df16-445">-ActiveView</span></span><br><span data-ttu-id="9df16-446">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9df16-446">
         -CompressedFile</span></span><br><span data-ttu-id="9df16-447">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-447">
         -DocumentEvents</span></span><br><span data-ttu-id="9df16-448">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9df16-448">
         - File</span></span><br><span data-ttu-id="9df16-449">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-449">
         -ImageCoercion</span></span><br><span data-ttu-id="9df16-450">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9df16-450">
         -PdfFile</span></span><br><span data-ttu-id="9df16-451">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9df16-451">
         - Selection</span></span><br><span data-ttu-id="9df16-452">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-452">
         - Settings</span></span><br><span data-ttu-id="9df16-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-453">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-454">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="9df16-454">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9df16-455">- 内容</span><span class="sxs-lookup"><span data-stu-id="9df16-455">- Content</span></span><br><span data-ttu-id="9df16-456">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-456">
         - Taskpane</span></span><br><span data-ttu-id="9df16-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9df16-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9df16-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9df16-459">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9df16-459">-ActiveView</span></span><br><span data-ttu-id="9df16-460">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9df16-460">
         -CompressedFile</span></span><br><span data-ttu-id="9df16-461">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-461">
         -DocumentEvents</span></span><br><span data-ttu-id="9df16-462">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9df16-462">
         - File</span></span><br><span data-ttu-id="9df16-463">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-463">
         -ImageCoercion</span></span><br><span data-ttu-id="9df16-464">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9df16-464">
         -PdfFile</span></span><br><span data-ttu-id="9df16-465">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9df16-465">
         - Selection</span></span><br><span data-ttu-id="9df16-466">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-466">
         - Settings</span></span><br><span data-ttu-id="9df16-467">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-467">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-468">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="9df16-468">Office for iOS</span></span></td>
    <td> <span data-ttu-id="9df16-469">- 内容</span><span class="sxs-lookup"><span data-stu-id="9df16-469">- Content</span></span><br><span data-ttu-id="9df16-470">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-470">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="9df16-471">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-471">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="9df16-472">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9df16-472">-ActiveView</span></span><br><span data-ttu-id="9df16-473">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9df16-473">
         -CompressedFile</span></span><br><span data-ttu-id="9df16-474">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-474">
         -DocumentEvents</span></span><br><span data-ttu-id="9df16-475">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9df16-475">
         - File</span></span><br><span data-ttu-id="9df16-476">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9df16-476">
         -PdfFile</span></span><br><span data-ttu-id="9df16-477">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9df16-477">
         - Selection</span></span><br><span data-ttu-id="9df16-478">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-478">
         - Settings</span></span><br><span data-ttu-id="9df16-479">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-479">
         -TextCoercion</span></span><br><span data-ttu-id="9df16-480">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-480">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-481">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="9df16-481">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="9df16-482">- 内容</span><span class="sxs-lookup"><span data-stu-id="9df16-482">- Content</span></span><br><span data-ttu-id="9df16-483">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-483">
         - Taskpane</span></span><br><span data-ttu-id="9df16-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9df16-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9df16-485">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-485">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9df16-486">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9df16-486">-ActiveView</span></span><br><span data-ttu-id="9df16-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9df16-487">
         -CompressedFile</span></span><br><span data-ttu-id="9df16-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-488">
         -DocumentEvents</span></span><br><span data-ttu-id="9df16-489">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9df16-489">
         - File</span></span><br><span data-ttu-id="9df16-490">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-490">
         -ImageCoercion</span></span><br><span data-ttu-id="9df16-491">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9df16-491">
         -PdfFile</span></span><br><span data-ttu-id="9df16-492">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9df16-492">
         - Selection</span></span><br><span data-ttu-id="9df16-493">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-493">
         - Settings</span></span><br><span data-ttu-id="9df16-494">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-494">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="9df16-495">OneNote</span><span class="sxs-lookup"><span data-stu-id="9df16-495">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9df16-496">平台</span><span class="sxs-lookup"><span data-stu-id="9df16-496">Platform</span></span></th>
    <th><span data-ttu-id="9df16-497">扩展点</span><span class="sxs-lookup"><span data-stu-id="9df16-497">Extension points</span></span></th>
    <th><span data-ttu-id="9df16-498">API 要求集</span><span class="sxs-lookup"><span data-stu-id="9df16-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="9df16-499"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="9df16-499"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="9df16-500">Office Online</span><span class="sxs-lookup"><span data-stu-id="9df16-500">Office Online</span></span></td>
    <td> <span data-ttu-id="9df16-501">- 内容</span><span class="sxs-lookup"><span data-stu-id="9df16-501">- Content</span></span><br><span data-ttu-id="9df16-502">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9df16-502">
         - Taskpane</span></span><br><span data-ttu-id="9df16-503">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9df16-503">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9df16-504">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-504">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="9df16-505">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9df16-505">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9df16-506">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9df16-506">-DocumentEvents</span></span><br><span data-ttu-id="9df16-507">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-507">
         -HtmlCoercion</span></span><br><span data-ttu-id="9df16-508">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-508">
         -ImageCoercion</span></span><br><span data-ttu-id="9df16-509">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="9df16-509">
         - Settings</span></span><br><span data-ttu-id="9df16-510">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9df16-510">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="9df16-511">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9df16-511">See also</span></span>

- [<span data-ttu-id="9df16-512">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="9df16-512">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="9df16-513">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="9df16-513">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="9df16-514">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="9df16-514">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="9df16-515">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="9df16-515">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)

---
title: Office 加载项主机和平台可用性
description: Excel、Word、Outlook、PowerPoint 和 OneNote 支持的要求集。
ms.date: 10/03/2018
ms.openlocfilehash: 6f7b5b565773457e6cd8a9eee69eb304784a29a9
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459313"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="72815-103">Office 加载项主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="72815-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="72815-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表包含每个 Office 应用程序目前支持的可用平台、扩展点、API 要求集和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="72815-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="72815-p102">如果表格单元格内有星号 (\*)，表示我们正在完善它。有关 Project 或 Access 要求集，请参阅 [Office 通用要求集](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="72815-p102">If a table cell contains an asterisk ( \* ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="72815-p103">通过 MSI 安装的 Office 2016 的内部版本号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="72815-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="72815-110">Excel</span><span class="sxs-lookup"><span data-stu-id="72815-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="72815-111">平台</span><span class="sxs-lookup"><span data-stu-id="72815-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="72815-112">扩展点</span><span class="sxs-lookup"><span data-stu-id="72815-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="72815-113">API 要求集</span><span class="sxs-lookup"><span data-stu-id="72815-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="72815-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72815-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="72815-115">Office Online</span></span></td>
    <td> <span data-ttu-id="72815-116">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-116">- Taskpane</span></span><br><span data-ttu-id="72815-117">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72815-117">
        - Content</span></span><br><span data-ttu-id="72815-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="72815-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="72815-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72815-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72815-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72815-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72815-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72815-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72815-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72815-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72815-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72815-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72815-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="72815-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72815-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72815-127">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-127">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="72815-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72815-128">
        -BindingEvents</span></span><br><span data-ttu-id="72815-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-129">
        -CompressedFile</span></span><br><span data-ttu-id="72815-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-130">
        -DocumentEvents</span></span><br><span data-ttu-id="72815-131">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-131">
        - File</span></span><br><span data-ttu-id="72815-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72815-132">
        -MatrixBindings</span></span><br><span data-ttu-id="72815-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-133">
        -MatrixCoercion</span></span><br><span data-ttu-id="72815-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-134">
        - Selection</span></span><br><span data-ttu-id="72815-135">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="72815-135">
        - Settings</span></span><br><span data-ttu-id="72815-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72815-136">
        -TableBindings</span></span><br><span data-ttu-id="72815-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-137">
        -TableCoercion</span></span><br><span data-ttu-id="72815-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72815-138">
        -TextBindings</span></span><br><span data-ttu-id="72815-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-139">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-140">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="72815-140">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="72815-141">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-141">
        - Taskpane</span></span><br><span data-ttu-id="72815-142">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72815-142">
        - Content</span></span></td>
    <td>  <span data-ttu-id="72815-143">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-143">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="72815-144">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72815-144">
        -BindingEvents</span></span><br><span data-ttu-id="72815-145">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-145">
        -CompressedFile</span></span><br><span data-ttu-id="72815-146">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-146">
        -DocumentEvents</span></span><br><span data-ttu-id="72815-147">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-147">
        - File</span></span><br><span data-ttu-id="72815-148">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-148">
        -ImageCoercion</span></span><br><span data-ttu-id="72815-149">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72815-149">
        -MatrixBindings</span></span><br><span data-ttu-id="72815-150">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-150">
        -MatrixCoercion</span></span><br><span data-ttu-id="72815-151">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-151">
        - Selection</span></span><br><span data-ttu-id="72815-152">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="72815-152">
        - Settings</span></span><br><span data-ttu-id="72815-153">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72815-153">
        -TableBindings</span></span><br><span data-ttu-id="72815-154">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-154">
        -TableCoercion</span></span><br><span data-ttu-id="72815-155">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72815-155">
        -TextBindings</span></span><br><span data-ttu-id="72815-156">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-156">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-157">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="72815-157">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="72815-158">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-158">- Taskpane</span></span><br><span data-ttu-id="72815-159">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72815-159">
        - Content</span></span><br><span data-ttu-id="72815-160">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-160">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="72815-161">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-161">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72815-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72815-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72815-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72815-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72815-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72815-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72815-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72815-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72815-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72815-167">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="72815-168">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72815-168">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72815-169">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-169">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="72815-170">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72815-170">-BindingEvents</span></span><br><span data-ttu-id="72815-171">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-171">
        -CompressedFile</span></span><br><span data-ttu-id="72815-172">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-172">
        -DocumentEvents</span></span><br><span data-ttu-id="72815-173">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-173">
        - File</span></span><br><span data-ttu-id="72815-174">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-174">
        -ImageCoercion</span></span><br><span data-ttu-id="72815-175">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72815-175">
        -MatrixBindings</span></span><br><span data-ttu-id="72815-176">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-176">
        -MatrixCoercion</span></span><br><span data-ttu-id="72815-177">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-177">
        - Selection</span></span><br><span data-ttu-id="72815-178">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="72815-178">
        - Settings</span></span><br><span data-ttu-id="72815-179">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72815-179">
        -TableBindings</span></span><br><span data-ttu-id="72815-180">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-180">
        -TableCoercion</span></span><br><span data-ttu-id="72815-181">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72815-181">
        -TextBindings</span></span><br><span data-ttu-id="72815-182">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-182">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-183">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="72815-183">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="72815-184">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-184">- Taskpane</span></span><br><span data-ttu-id="72815-185">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72815-185">
        - Content</span></span><br><span data-ttu-id="72815-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="72815-187">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-187">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72815-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72815-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72815-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72815-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72815-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72815-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72815-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72815-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72815-193">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72815-193">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="72815-194">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72815-194">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72815-195">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-195">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="72815-196">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72815-196">-BindingEvents</span></span><br><span data-ttu-id="72815-197">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-197">
        -CompressedFile</span></span><br><span data-ttu-id="72815-198">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-198">
        -DocumentEvents</span></span><br><span data-ttu-id="72815-199">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-199">
        - File</span></span><br><span data-ttu-id="72815-200">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-200">
        -ImageCoercion</span></span><br><span data-ttu-id="72815-201">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72815-201">
        -MatrixBindings</span></span><br><span data-ttu-id="72815-202">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-202">
        -MatrixCoercion</span></span><br><span data-ttu-id="72815-203">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-203">
        - Selection</span></span><br><span data-ttu-id="72815-204">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="72815-204">
        - Settings</span></span><br><span data-ttu-id="72815-205">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72815-205">
        -TableBindings</span></span><br><span data-ttu-id="72815-206">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-206">
        -TableCoercion</span></span><br><span data-ttu-id="72815-207">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72815-207">
        -TextBindings</span></span><br><span data-ttu-id="72815-208">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-208">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-209">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="72815-209">Office for iOS</span></span></td>
    <td><span data-ttu-id="72815-210">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-210">- Taskpane</span></span><br><span data-ttu-id="72815-211">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72815-211">
        - Content</span></span></td>
    <td><span data-ttu-id="72815-212">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-212">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72815-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72815-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72815-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72815-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72815-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72815-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72815-217">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72815-217">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72815-218">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72815-218">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="72815-219">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72815-219">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72815-220">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-220">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="72815-221">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72815-221">-BindingEvents</span></span><br><span data-ttu-id="72815-222">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-222">
        -CompressedFile</span></span><br><span data-ttu-id="72815-223">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-223">
        -DocumentEvents</span></span><br><span data-ttu-id="72815-224">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-224">
        - File</span></span><br><span data-ttu-id="72815-225">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-225">
        -ImageCoercion</span></span><br><span data-ttu-id="72815-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72815-226">
        -MatrixBindings</span></span><br><span data-ttu-id="72815-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-227">
        -MatrixCoercion</span></span><br><span data-ttu-id="72815-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-228">
        - Selection</span></span><br><span data-ttu-id="72815-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72815-229">
        - Settings</span></span><br><span data-ttu-id="72815-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72815-230">
        -TableBindings</span></span><br><span data-ttu-id="72815-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-231">
        -TableCoercion</span></span><br><span data-ttu-id="72815-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72815-232">
        -TextBindings</span></span><br><span data-ttu-id="72815-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-233">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-234">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="72815-234">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="72815-235">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-235">- Taskpane</span></span><br><span data-ttu-id="72815-236">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72815-236">
        - Content</span></span><br><span data-ttu-id="72815-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="72815-238">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-238">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72815-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72815-240">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-240">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72815-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72815-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72815-242">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72815-242">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72815-243">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72815-243">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72815-244">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72815-244">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="72815-245">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72815-245">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72815-246">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-246">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="72815-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72815-247">-BindingEvents</span></span><br><span data-ttu-id="72815-248">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-248">
        -CompressedFile</span></span><br><span data-ttu-id="72815-249">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-249">
        -DocumentEvents</span></span><br><span data-ttu-id="72815-250">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-250">
        - File</span></span><br><span data-ttu-id="72815-251">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-251">
        -ImageCoercion</span></span><br><span data-ttu-id="72815-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72815-252">
        -MatrixBindings</span></span><br><span data-ttu-id="72815-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-253">
        -MatrixCoercion</span></span><br><span data-ttu-id="72815-254">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-254">
        -PdfFile</span></span><br><span data-ttu-id="72815-255">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-255">
        - Selection</span></span><br><span data-ttu-id="72815-256">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="72815-256">
        - Settings</span></span><br><span data-ttu-id="72815-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72815-257">
        -TableBindings</span></span><br><span data-ttu-id="72815-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-258">
        -TableCoercion</span></span><br><span data-ttu-id="72815-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72815-259">
        -TextBindings</span></span><br><span data-ttu-id="72815-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-260">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-261">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="72815-261">Office for Mac</span></span></td>
    <td><span data-ttu-id="72815-262">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-262">- Taskpane</span></span><br><span data-ttu-id="72815-263">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72815-263">
        - Content</span></span><br><span data-ttu-id="72815-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="72815-265">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-265">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72815-266">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-266">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72815-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72815-268">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72815-268">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72815-269">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72815-269">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72815-270">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72815-270">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72815-271">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72815-271">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="72815-272">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72815-272">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72815-273">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-273">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="72815-274">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72815-274">-BindingEvents</span></span><br><span data-ttu-id="72815-275">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-275">
        -CompressedFile</span></span><br><span data-ttu-id="72815-276">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-276">
        -DocumentEvents</span></span><br><span data-ttu-id="72815-277">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-277">
        - File</span></span><br><span data-ttu-id="72815-278">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-278">
        -ImageCoercion</span></span><br><span data-ttu-id="72815-279">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72815-279">
        -MatrixBindings</span></span><br><span data-ttu-id="72815-280">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-280">
        -MatrixCoercion</span></span><br><span data-ttu-id="72815-281">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-281">
        -PdfFile</span></span><br><span data-ttu-id="72815-282">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-282">
        - Selection</span></span><br><span data-ttu-id="72815-283">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="72815-283">
        - Settings</span></span><br><span data-ttu-id="72815-284">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72815-284">
        -TableBindings</span></span><br><span data-ttu-id="72815-285">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-285">
        -TableCoercion</span></span><br><span data-ttu-id="72815-286">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72815-286">
        -TextBindings</span></span><br><span data-ttu-id="72815-287">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-287">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="72815-288">Outlook</span><span class="sxs-lookup"><span data-stu-id="72815-288">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="72815-289">平台</span><span class="sxs-lookup"><span data-stu-id="72815-289">Platform</span></span></th>
    <th><span data-ttu-id="72815-290">扩展点</span><span class="sxs-lookup"><span data-stu-id="72815-290">Extension points</span></span></th>
    <th><span data-ttu-id="72815-291">API 要求集</span><span class="sxs-lookup"><span data-stu-id="72815-291">API requirement sets</span></span></th>
    <th><span data-ttu-id="72815-292"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72815-292"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-293">Office Online</span><span class="sxs-lookup"><span data-stu-id="72815-293">Office Online</span></span></td>
    <td> <span data-ttu-id="72815-294">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72815-294">- Mail Read</span></span><br><span data-ttu-id="72815-295">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="72815-295">
      - Mail Compose</span></span><br><span data-ttu-id="72815-296">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-296">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-297">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-297">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72815-298">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-298">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72815-299">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-299">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72815-300">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72815-300">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72815-301">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72815-301">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72815-302">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72815-302">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="72815-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72815-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="72815-304">不适用</span><span class="sxs-lookup"><span data-stu-id="72815-304">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-305">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="72815-305">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="72815-306">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72815-306">- Mail Read</span></span><br><span data-ttu-id="72815-307">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="72815-307">
      - Mail Compose</span></span><br><span data-ttu-id="72815-308">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-308">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-309">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-309">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72815-310">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-310">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72815-311">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-311">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72815-312">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72815-312">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="72815-313">不适用</span><span class="sxs-lookup"><span data-stu-id="72815-313">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-314">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="72815-314">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="72815-315">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72815-315">- Mail Read</span></span><br><span data-ttu-id="72815-316">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="72815-316">
      - Mail Compose</span></span><br><span data-ttu-id="72815-317">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-317">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="72815-318">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="72815-318">
      - Modules</span></span></td>
    <td> <span data-ttu-id="72815-319">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-319">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72815-320">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-320">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72815-321">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-321">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72815-322">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72815-322">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72815-323">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72815-323">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72815-324">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72815-324">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="72815-325">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72815-325">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="72815-326">不适用</span><span class="sxs-lookup"><span data-stu-id="72815-326">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-327">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="72815-327">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="72815-328">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72815-328">- Mail Read</span></span><br><span data-ttu-id="72815-329">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="72815-329">
      - Mail Compose</span></span><br><span data-ttu-id="72815-330">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-330">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="72815-331">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="72815-331">
      - Modules</span></span></td>
    <td> <span data-ttu-id="72815-332">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-332">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72815-333">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-333">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72815-334">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-334">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72815-335">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72815-335">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72815-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72815-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72815-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72815-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="72815-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72815-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="72815-339">不适用</span><span class="sxs-lookup"><span data-stu-id="72815-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-340">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="72815-340">Office for iOS</span></span></td>
    <td> <span data-ttu-id="72815-341">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72815-341">- Mail Read</span></span><br><span data-ttu-id="72815-342">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-342">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-343">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-343">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72815-344">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-344">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72815-345">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-345">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72815-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72815-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72815-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72815-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="72815-348">不适用</span><span class="sxs-lookup"><span data-stu-id="72815-348">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-349">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="72815-349">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="72815-350">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72815-350">- Mail Read</span></span><br><span data-ttu-id="72815-351">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="72815-351">
      - Mail Compose</span></span><br><span data-ttu-id="72815-352">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-352">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-353">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-353">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72815-354">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-354">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72815-355">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-355">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72815-356">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72815-356">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72815-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72815-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72815-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72815-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="72815-359">不适用</span><span class="sxs-lookup"><span data-stu-id="72815-359">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-360">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="72815-360">Office for Mac</span></span></td>
    <td> <span data-ttu-id="72815-361">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72815-361">- Mail Read</span></span><br><span data-ttu-id="72815-362">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="72815-362">
      - Mail Compose</span></span><br><span data-ttu-id="72815-363">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-363">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-364">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-364">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72815-365">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-365">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72815-366">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-366">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72815-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72815-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72815-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72815-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72815-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72815-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="72815-370">不适用</span><span class="sxs-lookup"><span data-stu-id="72815-370">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-371">Office for Android</span><span class="sxs-lookup"><span data-stu-id="72815-371">Office for Android</span></span></td>
    <td> <span data-ttu-id="72815-372">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72815-372">- Mail Read</span></span><br><span data-ttu-id="72815-373">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-373">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-374">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-374">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72815-375">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-375">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72815-376">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-376">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72815-377">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72815-377">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72815-378">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72815-378">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="72815-379">不适用</span><span class="sxs-lookup"><span data-stu-id="72815-379">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="72815-380">Word</span><span class="sxs-lookup"><span data-stu-id="72815-380">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="72815-381">平台</span><span class="sxs-lookup"><span data-stu-id="72815-381">Platform</span></span></th>
    <th><span data-ttu-id="72815-382">扩展点</span><span class="sxs-lookup"><span data-stu-id="72815-382">Extension points</span></span></th>
    <th><span data-ttu-id="72815-383">API 要求集</span><span class="sxs-lookup"><span data-stu-id="72815-383">API requirement sets</span></span></th>
    <th><span data-ttu-id="72815-384"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72815-384"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="72815-385">Office Online</span><span class="sxs-lookup"><span data-stu-id="72815-385">Office Online</span></span></td>
    <td> <span data-ttu-id="72815-386">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-386">- Taskpane</span></span><br><span data-ttu-id="72815-387">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-387">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-388">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-388">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="72815-389">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-389">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="72815-390">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-390">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="72815-391">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-391">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72815-392">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72815-392">-BindingEvents</span></span><br><span data-ttu-id="72815-393">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72815-393">
         -</span></span><br><span data-ttu-id="72815-394">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-394">
         -DocumentEvents</span></span><br><span data-ttu-id="72815-395">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-395">
         - File</span></span><br><span data-ttu-id="72815-396">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-396">
         -HtmlCoercion</span></span><br><span data-ttu-id="72815-397">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-397">
         -ImageCoercion</span></span><br><span data-ttu-id="72815-398">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72815-398">
         -MatrixBindings</span></span><br><span data-ttu-id="72815-399">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-399">
         -MatrixCoercion</span></span><br><span data-ttu-id="72815-400">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-400">
         -OoxmlCoercion</span></span><br><span data-ttu-id="72815-401">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-401">
         -PdfFile</span></span><br><span data-ttu-id="72815-402">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-402">
         - Selection</span></span><br><span data-ttu-id="72815-403">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="72815-403">
         - Settings</span></span><br><span data-ttu-id="72815-404">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72815-404">
         -TableBindings</span></span><br><span data-ttu-id="72815-405">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-405">
         -TableCoercion</span></span><br><span data-ttu-id="72815-406">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72815-406">
         -TextBindings</span></span><br><span data-ttu-id="72815-407">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-407">
         -TextCoercion</span></span><br><span data-ttu-id="72815-408">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72815-408">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-409">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="72815-409">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="72815-410">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-410">- Taskpane</span></span></td>
    <td> <span data-ttu-id="72815-411">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-411">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72815-412">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72815-412">-BindingEvents</span></span><br><span data-ttu-id="72815-413">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-413">
         -CompressedFile</span></span><br><span data-ttu-id="72815-414">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72815-414">
         -</span></span><br><span data-ttu-id="72815-415">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-415">
         -DocumentEvents</span></span><br><span data-ttu-id="72815-416">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-416">
         - File</span></span><br><span data-ttu-id="72815-417">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-417">
         -HtmlCoercion</span></span><br><span data-ttu-id="72815-418">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-418">
         -ImageCoercion</span></span><br><span data-ttu-id="72815-419">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72815-419">
         -MatrixBindings</span></span><br><span data-ttu-id="72815-420">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-420">
         -MatrixCoercion</span></span><br><span data-ttu-id="72815-421">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-421">
         -OoxmlCoercion</span></span><br><span data-ttu-id="72815-422">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-422">
         -PdfFile</span></span><br><span data-ttu-id="72815-423">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-423">
         - Selection</span></span><br><span data-ttu-id="72815-424">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="72815-424">
         - Settings</span></span><br><span data-ttu-id="72815-425">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72815-425">
         -TableBindings</span></span><br><span data-ttu-id="72815-426">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-426">
         -TableCoercion</span></span><br><span data-ttu-id="72815-427">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72815-427">
         -TextBindings</span></span><br><span data-ttu-id="72815-428">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-428">
         -TextCoercion</span></span><br><span data-ttu-id="72815-429">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72815-429">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-430">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="72815-430">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="72815-431">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-431">- Taskpane</span></span><br><span data-ttu-id="72815-432">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-432">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-433">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-433">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="72815-434">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-434">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="72815-435">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-435">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="72815-436">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-436">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72815-437">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72815-437">-BindingEvents</span></span><br><span data-ttu-id="72815-438">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-438">
         -CompressedFile</span></span><br><span data-ttu-id="72815-439">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72815-439">
         -</span></span><br><span data-ttu-id="72815-440">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-440">
         -DocumentEvents</span></span><br><span data-ttu-id="72815-441">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-441">
         - File</span></span><br><span data-ttu-id="72815-442">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-442">
         -HtmlCoercion</span></span><br><span data-ttu-id="72815-443">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-443">
         -ImageCoercion</span></span><br><span data-ttu-id="72815-444">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72815-444">
         -MatrixBindings</span></span><br><span data-ttu-id="72815-445">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-445">
         -MatrixCoercion</span></span><br><span data-ttu-id="72815-446">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-446">
         -OoxmlCoercion</span></span><br><span data-ttu-id="72815-447">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-447">
         -PdfFile</span></span><br><span data-ttu-id="72815-448">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-448">
         - Selection</span></span><br><span data-ttu-id="72815-449">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="72815-449">
         - Settings</span></span><br><span data-ttu-id="72815-450">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72815-450">
         -TableBindings</span></span><br><span data-ttu-id="72815-451">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-451">
         -TableCoercion</span></span><br><span data-ttu-id="72815-452">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72815-452">
         -TextBindings</span></span><br><span data-ttu-id="72815-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-453">
         -TextCoercion</span></span><br><span data-ttu-id="72815-454">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72815-454">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-455">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="72815-455">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="72815-456">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-456">- Taskpane</span></span><br><span data-ttu-id="72815-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="72815-459">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-459">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="72815-460">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-460">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="72815-461">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-461">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72815-462">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72815-462">-BindingEvents</span></span><br><span data-ttu-id="72815-463">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-463">
         -CompressedFile</span></span><br><span data-ttu-id="72815-464">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72815-464">
         -</span></span><br><span data-ttu-id="72815-465">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-465">
         -DocumentEvents</span></span><br><span data-ttu-id="72815-466">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-466">
         - File</span></span><br><span data-ttu-id="72815-467">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-467">
         -HtmlCoercion</span></span><br><span data-ttu-id="72815-468">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-468">
         -ImageCoercion</span></span><br><span data-ttu-id="72815-469">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72815-469">
         -MatrixBindings</span></span><br><span data-ttu-id="72815-470">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-470">
         -MatrixCoercion</span></span><br><span data-ttu-id="72815-471">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-471">
         -OoxmlCoercion</span></span><br><span data-ttu-id="72815-472">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-472">
         -PdfFile</span></span><br><span data-ttu-id="72815-473">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-473">
         - Selection</span></span><br><span data-ttu-id="72815-474">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="72815-474">
         - Settings</span></span><br><span data-ttu-id="72815-475">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72815-475">
         -TableBindings</span></span><br><span data-ttu-id="72815-476">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-476">
         -TableCoercion</span></span><br><span data-ttu-id="72815-477">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72815-477">
         -TextBindings</span></span><br><span data-ttu-id="72815-478">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-478">
         -TextCoercion</span></span><br><span data-ttu-id="72815-479">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72815-479">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-480">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="72815-480">Office for iOS</span></span></td>
    <td> <span data-ttu-id="72815-481">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-481">- Taskpane</span></span></td>
    <td> <span data-ttu-id="72815-482">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-482">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="72815-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="72815-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="72815-485">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="72815-485">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="72815-486">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72815-486">-BindingEvents</span></span><br><span data-ttu-id="72815-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-487">
         -CompressedFile</span></span><br><span data-ttu-id="72815-488">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72815-488">
         -</span></span><br><span data-ttu-id="72815-489">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-489">
         -DocumentEvents</span></span><br><span data-ttu-id="72815-490">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-490">
         - File</span></span><br><span data-ttu-id="72815-491">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-491">
         -HtmlCoercion</span></span><br><span data-ttu-id="72815-492">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-492">
         -ImageCoercion</span></span><br><span data-ttu-id="72815-493">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72815-493">
         -MatrixBindings</span></span><br><span data-ttu-id="72815-494">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-494">
         -MatrixCoercion</span></span><br><span data-ttu-id="72815-495">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-495">
         -OoxmlCoercion</span></span><br><span data-ttu-id="72815-496">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-496">
         -PdfFile</span></span><br><span data-ttu-id="72815-497">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-497">
         - Selection</span></span><br><span data-ttu-id="72815-498">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="72815-498">
         - Settings</span></span><br><span data-ttu-id="72815-499">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72815-499">
         -TableBindings</span></span><br><span data-ttu-id="72815-500">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-500">
         -TableCoercion</span></span><br><span data-ttu-id="72815-501">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72815-501">
         -TextBindings</span></span><br><span data-ttu-id="72815-502">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-502">
         -TextCoercion</span></span><br><span data-ttu-id="72815-503">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72815-503">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-504">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="72815-504">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="72815-505">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-505">- Taskpane</span></span><br><span data-ttu-id="72815-506">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-506">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-507">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-507">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="72815-508">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-508">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="72815-509">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-509">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="72815-510">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="72815-510">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="72815-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72815-511">-BindingEvents</span></span><br><span data-ttu-id="72815-512">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-512">
         -CompressedFile</span></span><br><span data-ttu-id="72815-513">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72815-513">
         -</span></span><br><span data-ttu-id="72815-514">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-514">
         -DocumentEvents</span></span><br><span data-ttu-id="72815-515">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-515">
         - File</span></span><br><span data-ttu-id="72815-516">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-516">
         -HtmlCoercion</span></span><br><span data-ttu-id="72815-517">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-517">
         -ImageCoercion</span></span><br><span data-ttu-id="72815-518">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72815-518">
         -MatrixBindings</span></span><br><span data-ttu-id="72815-519">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-519">
         -MatrixCoercion</span></span><br><span data-ttu-id="72815-520">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-520">
         -OoxmlCoercion</span></span><br><span data-ttu-id="72815-521">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-521">
         -PdfFile</span></span><br><span data-ttu-id="72815-522">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-522">
         - Selection</span></span><br><span data-ttu-id="72815-523">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="72815-523">
         - Settings</span></span><br><span data-ttu-id="72815-524">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72815-524">
         -TableBindings</span></span><br><span data-ttu-id="72815-525">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-525">
         -TableCoercion</span></span><br><span data-ttu-id="72815-526">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72815-526">
         -TextBindings</span></span><br><span data-ttu-id="72815-527">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-527">
         -TextCoercion</span></span><br><span data-ttu-id="72815-528">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72815-528">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-529">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="72815-529">Office for Mac</span></span></td>
    <td> <span data-ttu-id="72815-530">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-530">- Taskpane</span></span><br><span data-ttu-id="72815-531">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-531">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-532">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-532">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="72815-533">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72815-533">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="72815-534">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72815-534">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="72815-535">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="72815-535">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="72815-536">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72815-536">-BindingEvents</span></span><br><span data-ttu-id="72815-537">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-537">
         -CompressedFile</span></span><br><span data-ttu-id="72815-538">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72815-538">
         -</span></span><br><span data-ttu-id="72815-539">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-539">
         -DocumentEvents</span></span><br><span data-ttu-id="72815-540">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-540">
         - File</span></span><br><span data-ttu-id="72815-541">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-541">
         -HtmlCoercion</span></span><br><span data-ttu-id="72815-542">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-542">
         -ImageCoercion</span></span><br><span data-ttu-id="72815-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72815-543">
         -MatrixBindings</span></span><br><span data-ttu-id="72815-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-544">
         -MatrixCoercion</span></span><br><span data-ttu-id="72815-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-545">
         -OoxmlCoercion</span></span><br><span data-ttu-id="72815-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-546">
         -PdfFile</span></span><br><span data-ttu-id="72815-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-547">
         - Selection</span></span><br><span data-ttu-id="72815-548">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="72815-548">
         - Settings</span></span><br><span data-ttu-id="72815-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72815-549">
         -TableBindings</span></span><br><span data-ttu-id="72815-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-550">
         -TableCoercion</span></span><br><span data-ttu-id="72815-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72815-551">
         -TextBindings</span></span><br><span data-ttu-id="72815-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-552">
         -TextCoercion</span></span><br><span data-ttu-id="72815-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72815-553">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="72815-554">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="72815-554">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="72815-555">平台</span><span class="sxs-lookup"><span data-stu-id="72815-555">Platform</span></span></th>
    <th><span data-ttu-id="72815-556">扩展点</span><span class="sxs-lookup"><span data-stu-id="72815-556">Extension points</span></span></th>
    <th><span data-ttu-id="72815-557">API 要求集</span><span class="sxs-lookup"><span data-stu-id="72815-557">API requirement sets</span></span></th>
    <th><span data-ttu-id="72815-558"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72815-558"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="72815-559">Office Online</span><span class="sxs-lookup"><span data-stu-id="72815-559">Office Online</span></span></td>
    <td> <span data-ttu-id="72815-560">- 内容</span><span class="sxs-lookup"><span data-stu-id="72815-560">- Content</span></span><br><span data-ttu-id="72815-561">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-561">
         - Taskpane</span></span><br><span data-ttu-id="72815-562">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-562">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-563">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-563">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72815-564">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72815-564">-ActiveView</span></span><br><span data-ttu-id="72815-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-565">
         -CompressedFile</span></span><br><span data-ttu-id="72815-566">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-566">
         -DocumentEvents</span></span><br><span data-ttu-id="72815-567">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-567">
         - File</span></span><br><span data-ttu-id="72815-568">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-568">
         -ImageCoercion</span></span><br><span data-ttu-id="72815-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-569">
         -PdfFile</span></span><br><span data-ttu-id="72815-570">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-570">
         - Selection</span></span><br><span data-ttu-id="72815-571">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72815-571">
         - Settings</span></span><br><span data-ttu-id="72815-572">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-572">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-573">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="72815-573">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="72815-574">- 内容</span><span class="sxs-lookup"><span data-stu-id="72815-574">- Content</span></span><br><span data-ttu-id="72815-575">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-575">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="72815-576">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="72815-576">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="72815-577">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72815-577">-ActiveView</span></span><br><span data-ttu-id="72815-578">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-578">
         -CompressedFile</span></span><br><span data-ttu-id="72815-579">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-579">
         -DocumentEvents</span></span><br><span data-ttu-id="72815-580">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-580">
         - File</span></span><br><span data-ttu-id="72815-581">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-581">
         -ImageCoercion</span></span><br><span data-ttu-id="72815-582">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-582">
         -PdfFile</span></span><br><span data-ttu-id="72815-583">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-583">
         - Selection</span></span><br><span data-ttu-id="72815-584">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72815-584">
         - Settings</span></span><br><span data-ttu-id="72815-585">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-585">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-586">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="72815-586">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="72815-587">- 内容</span><span class="sxs-lookup"><span data-stu-id="72815-587">- Content</span></span><br><span data-ttu-id="72815-588">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-588">
         - Taskpane</span></span><br><span data-ttu-id="72815-589">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-589">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-590">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-590">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72815-591">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72815-591">-ActiveView</span></span><br><span data-ttu-id="72815-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-592">
         -CompressedFile</span></span><br><span data-ttu-id="72815-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-593">
         -DocumentEvents</span></span><br><span data-ttu-id="72815-594">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-594">
         - File</span></span><br><span data-ttu-id="72815-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-595">
         -ImageCoercion</span></span><br><span data-ttu-id="72815-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-596">
         -PdfFile</span></span><br><span data-ttu-id="72815-597">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-597">
         - Selection</span></span><br><span data-ttu-id="72815-598">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72815-598">
         - Settings</span></span><br><span data-ttu-id="72815-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-599">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-600">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="72815-600">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="72815-601">- 内容</span><span class="sxs-lookup"><span data-stu-id="72815-601">- Content</span></span><br><span data-ttu-id="72815-602">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-602">
         - Taskpane</span></span><br><span data-ttu-id="72815-603">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-603">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-604">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-604">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72815-605">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72815-605">-ActiveView</span></span><br><span data-ttu-id="72815-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-606">
         -CompressedFile</span></span><br><span data-ttu-id="72815-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-607">
         -DocumentEvents</span></span><br><span data-ttu-id="72815-608">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-608">
         - File</span></span><br><span data-ttu-id="72815-609">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-609">
         -ImageCoercion</span></span><br><span data-ttu-id="72815-610">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-610">
         -PdfFile</span></span><br><span data-ttu-id="72815-611">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-611">
         - Selection</span></span><br><span data-ttu-id="72815-612">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72815-612">
         - Settings</span></span><br><span data-ttu-id="72815-613">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-613">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-614">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="72815-614">Office for iOS</span></span></td>
    <td> <span data-ttu-id="72815-615">- 内容</span><span class="sxs-lookup"><span data-stu-id="72815-615">- Content</span></span><br><span data-ttu-id="72815-616">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-616">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="72815-617">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-617">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="72815-618">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72815-618">-ActiveView</span></span><br><span data-ttu-id="72815-619">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-619">
         -CompressedFile</span></span><br><span data-ttu-id="72815-620">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-620">
         -DocumentEvents</span></span><br><span data-ttu-id="72815-621">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-621">
         - File</span></span><br><span data-ttu-id="72815-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-622">
         -PdfFile</span></span><br><span data-ttu-id="72815-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-623">
         - Selection</span></span><br><span data-ttu-id="72815-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72815-624">
         - Settings</span></span><br><span data-ttu-id="72815-625">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-625">
         -TextCoercion</span></span><br><span data-ttu-id="72815-626">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-626">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-627">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="72815-627">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="72815-628">- 内容</span><span class="sxs-lookup"><span data-stu-id="72815-628">- Content</span></span><br><span data-ttu-id="72815-629">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-629">
         - Taskpane</span></span><br><span data-ttu-id="72815-630">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-630">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-631">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-631">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72815-632">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72815-632">-ActiveView</span></span><br><span data-ttu-id="72815-633">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-633">
         -CompressedFile</span></span><br><span data-ttu-id="72815-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-634">
         -DocumentEvents</span></span><br><span data-ttu-id="72815-635">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-635">
         - File</span></span><br><span data-ttu-id="72815-636">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-636">
         -ImageCoercion</span></span><br><span data-ttu-id="72815-637">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-637">
         -PdfFile</span></span><br><span data-ttu-id="72815-638">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-638">
         - Selection</span></span><br><span data-ttu-id="72815-639">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72815-639">
         - Settings</span></span><br><span data-ttu-id="72815-640">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-640">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72815-641">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="72815-641">Office for Mac</span></span></td>
    <td> <span data-ttu-id="72815-642">- 内容</span><span class="sxs-lookup"><span data-stu-id="72815-642">- Content</span></span><br><span data-ttu-id="72815-643">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-643">
         - Taskpane</span></span><br><span data-ttu-id="72815-644">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-644">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-645">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-645">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72815-646">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72815-646">-ActiveView</span></span><br><span data-ttu-id="72815-647">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72815-647">
         -CompressedFile</span></span><br><span data-ttu-id="72815-648">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-648">
         -DocumentEvents</span></span><br><span data-ttu-id="72815-649">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="72815-649">
         - File</span></span><br><span data-ttu-id="72815-650">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-650">
         -ImageCoercion</span></span><br><span data-ttu-id="72815-651">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72815-651">
         -PdfFile</span></span><br><span data-ttu-id="72815-652">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72815-652">
         - Selection</span></span><br><span data-ttu-id="72815-653">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72815-653">
         - Settings</span></span><br><span data-ttu-id="72815-654">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-654">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="72815-655">OneNote</span><span class="sxs-lookup"><span data-stu-id="72815-655">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="72815-656">平台</span><span class="sxs-lookup"><span data-stu-id="72815-656">Platform</span></span></th>
    <th><span data-ttu-id="72815-657">扩展点</span><span class="sxs-lookup"><span data-stu-id="72815-657">Extension points</span></span></th>
    <th><span data-ttu-id="72815-658">API 要求集</span><span class="sxs-lookup"><span data-stu-id="72815-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="72815-659"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72815-659"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="72815-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="72815-660">Office Online</span></span></td>
    <td> <span data-ttu-id="72815-661">- 内容</span><span class="sxs-lookup"><span data-stu-id="72815-661">- Content</span></span><br><span data-ttu-id="72815-662">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72815-662">
         - Taskpane</span></span><br><span data-ttu-id="72815-663">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72815-663">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72815-664">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-664">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="72815-665">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72815-665">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72815-666">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72815-666">-DocumentEvents</span></span><br><span data-ttu-id="72815-667">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-667">
         -HtmlCoercion</span></span><br><span data-ttu-id="72815-668">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-668">
         -ImageCoercion</span></span><br><span data-ttu-id="72815-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72815-669">
         - Settings</span></span><br><span data-ttu-id="72815-670">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72815-670">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="72815-671">另请参阅</span><span class="sxs-lookup"><span data-stu-id="72815-671">See also</span></span>

- [<span data-ttu-id="72815-672">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="72815-672">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="72815-673">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="72815-673">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="72815-674">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="72815-674">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="72815-675">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="72815-675">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)

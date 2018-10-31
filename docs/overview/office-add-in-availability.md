---
title: Office 加载项主机和平台可用性
description: Excel、Word、Outlook、PowerPoint 和 OneNote 支持的要求集。
ms.date: 10/03/2018
ms.openlocfilehash: 39a80f322c282e29e6e8c4363f0c82522b33b75d
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579924"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="ce7e2-103">Office 加载项主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="ce7e2-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="ce7e2-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表包含每个 Office 应用程序目前支持的可用平台、扩展点、API 要求集和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="ce7e2-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="ce7e2-p102">如果表格单元格内有星号 (\*)，表示我们正在完善它。有关 Project 或 Access 要求集，请参阅 [Office 通用要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="ce7e2-p102">If a table cell contains an asterisk ( \* ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="ce7e2-p103">通过 MSI 安装的 Office 2016 的内部版本号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="ce7e2-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="ce7e2-110">Excel</span><span class="sxs-lookup"><span data-stu-id="ce7e2-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ce7e2-111">平台</span><span class="sxs-lookup"><span data-stu-id="ce7e2-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ce7e2-112">扩展点</span><span class="sxs-lookup"><span data-stu-id="ce7e2-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ce7e2-113">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ce7e2-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ce7e2-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="ce7e2-115">Office Online</span></span></td>
    <td> <span data-ttu-id="ce7e2-116">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-116">- Taskpane</span></span><br><span data-ttu-id="ce7e2-117">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ce7e2-117">
        - Content</span></span><br><span data-ttu-id="ce7e2-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="ce7e2-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ce7e2-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ce7e2-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ce7e2-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ce7e2-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ce7e2-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ce7e2-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ce7e2-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="ce7e2-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ce7e2-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ce7e2-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-128">
        -BindingEvents</span></span><br><span data-ttu-id="ce7e2-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-129">
        -CompressedFile</span></span><br><span data-ttu-id="ce7e2-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-130">
        -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-131">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-131">
        - File</span></span><br><span data-ttu-id="ce7e2-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-132">
        -MatrixBindings</span></span><br><span data-ttu-id="ce7e2-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-133">
        -MatrixCoercion</span></span><br><span data-ttu-id="ce7e2-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-134">
        - Selection</span></span><br><span data-ttu-id="ce7e2-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-135">
        - Settings</span></span><br><span data-ttu-id="ce7e2-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-136">
        -TableBindings</span></span><br><span data-ttu-id="ce7e2-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-137">
        -TableCoercion</span></span><br><span data-ttu-id="ce7e2-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-138">
        -TextBindings</span></span><br><span data-ttu-id="ce7e2-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-139">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-140">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="ce7e2-140">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="ce7e2-141">
        - Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-141">
        - Taskpane</span></span><br><span data-ttu-id="ce7e2-142">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ce7e2-142">
        - Content</span></span></td>
    <td>  <span data-ttu-id="ce7e2-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ce7e2-144">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-144">
        -BindingEvents</span></span><br><span data-ttu-id="ce7e2-145">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-145">
        -CompressedFile</span></span><br><span data-ttu-id="ce7e2-146">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-146">
        -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-147">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-147">
        - File</span></span><br><span data-ttu-id="ce7e2-148">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-148">
        -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-149">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-149">
        -MatrixBindings</span></span><br><span data-ttu-id="ce7e2-150">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-150">
        -MatrixCoercion</span></span><br><span data-ttu-id="ce7e2-151">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-151">
        - Selection</span></span><br><span data-ttu-id="ce7e2-152">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-152">
        - Settings</span></span><br><span data-ttu-id="ce7e2-153">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-153">
        -TableBindings</span></span><br><span data-ttu-id="ce7e2-154">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-154">
        -TableCoercion</span></span><br><span data-ttu-id="ce7e2-155">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-155">
        -TextBindings</span></span><br><span data-ttu-id="ce7e2-156">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-156">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-157">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="ce7e2-157">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="ce7e2-158">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-158">- Taskpane</span></span><br><span data-ttu-id="ce7e2-159">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ce7e2-159">
        - Content</span></span><br><span data-ttu-id="ce7e2-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ce7e2-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ce7e2-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ce7e2-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ce7e2-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ce7e2-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ce7e2-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ce7e2-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-167">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="ce7e2-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ce7e2-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ce7e2-170">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-170">-BindingEvents</span></span><br><span data-ttu-id="ce7e2-171">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-171">
        -CompressedFile</span></span><br><span data-ttu-id="ce7e2-172">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-172">
        -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-173">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-173">
        - File</span></span><br><span data-ttu-id="ce7e2-174">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-174">
        -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-175">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-175">
        -MatrixBindings</span></span><br><span data-ttu-id="ce7e2-176">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-176">
        -MatrixCoercion</span></span><br><span data-ttu-id="ce7e2-177">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-177">
        - Selection</span></span><br><span data-ttu-id="ce7e2-178">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-178">
        - Settings</span></span><br><span data-ttu-id="ce7e2-179">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-179">
        -TableBindings</span></span><br><span data-ttu-id="ce7e2-180">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-180">
        -TableCoercion</span></span><br><span data-ttu-id="ce7e2-181">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-181">
        -TextBindings</span></span><br><span data-ttu-id="ce7e2-182">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-182">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-183">Office for Windows</span><span class="sxs-lookup"><span data-stu-id="ce7e2-183">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="ce7e2-184">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-184">- Taskpane</span></span><br><span data-ttu-id="ce7e2-185">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ce7e2-185">
        - Content</span></span><br><span data-ttu-id="ce7e2-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ce7e2-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ce7e2-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ce7e2-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ce7e2-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ce7e2-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ce7e2-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ce7e2-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-193">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="ce7e2-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ce7e2-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ce7e2-196">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-196">-BindingEvents</span></span><br><span data-ttu-id="ce7e2-197">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-197">
        -CompressedFile</span></span><br><span data-ttu-id="ce7e2-198">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-198">
        -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-199">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-199">
        - File</span></span><br><span data-ttu-id="ce7e2-200">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-200">
        -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-201">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-201">
        -MatrixBindings</span></span><br><span data-ttu-id="ce7e2-202">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-202">
        -MatrixCoercion</span></span><br><span data-ttu-id="ce7e2-203">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-203">
        - Selection</span></span><br><span data-ttu-id="ce7e2-204">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-204">
        - Settings</span></span><br><span data-ttu-id="ce7e2-205">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-205">
        -TableBindings</span></span><br><span data-ttu-id="ce7e2-206">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-206">
        -TableCoercion</span></span><br><span data-ttu-id="ce7e2-207">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-207">
        -TextBindings</span></span><br><span data-ttu-id="ce7e2-208">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-208">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-209">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="ce7e2-209">Office for iOS</span></span></td>
    <td><span data-ttu-id="ce7e2-210">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-210">- Taskpane</span></span><br><span data-ttu-id="ce7e2-211">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ce7e2-211">
        - Content</span></span></td>
    <td><span data-ttu-id="ce7e2-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ce7e2-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ce7e2-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ce7e2-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ce7e2-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ce7e2-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ce7e2-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-218">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="ce7e2-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ce7e2-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ce7e2-221">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-221">-BindingEvents</span></span><br><span data-ttu-id="ce7e2-222">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-222">
        -CompressedFile</span></span><br><span data-ttu-id="ce7e2-223">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-223">
        -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-224">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-224">
        - File</span></span><br><span data-ttu-id="ce7e2-225">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-225">
        -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-226">
        -MatrixBindings</span></span><br><span data-ttu-id="ce7e2-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-227">
        -MatrixCoercion</span></span><br><span data-ttu-id="ce7e2-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-228">
        - Selection</span></span><br><span data-ttu-id="ce7e2-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-229">
        - Settings</span></span><br><span data-ttu-id="ce7e2-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-230">
        -TableBindings</span></span><br><span data-ttu-id="ce7e2-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-231">
        -TableCoercion</span></span><br><span data-ttu-id="ce7e2-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-232">
        -TextBindings</span></span><br><span data-ttu-id="ce7e2-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-233">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-234">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="ce7e2-234">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="ce7e2-235">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-235">- Taskpane</span></span><br><span data-ttu-id="ce7e2-236">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ce7e2-236">
        - Content</span></span><br><span data-ttu-id="ce7e2-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ce7e2-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ce7e2-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ce7e2-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ce7e2-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ce7e2-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ce7e2-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ce7e2-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-244">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="ce7e2-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ce7e2-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ce7e2-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-247">-BindingEvents</span></span><br><span data-ttu-id="ce7e2-248">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-248">
        -CompressedFile</span></span><br><span data-ttu-id="ce7e2-249">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-249">
        -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-250">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-250">
        - File</span></span><br><span data-ttu-id="ce7e2-251">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-251">
        -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-252">
        -MatrixBindings</span></span><br><span data-ttu-id="ce7e2-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-253">
        -MatrixCoercion</span></span><br><span data-ttu-id="ce7e2-254">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-254">
        -PdfFile</span></span><br><span data-ttu-id="ce7e2-255">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-255">
        - Selection</span></span><br><span data-ttu-id="ce7e2-256">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-256">
        - Settings</span></span><br><span data-ttu-id="ce7e2-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-257">
        -TableBindings</span></span><br><span data-ttu-id="ce7e2-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-258">
        -TableCoercion</span></span><br><span data-ttu-id="ce7e2-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-259">
        -TextBindings</span></span><br><span data-ttu-id="ce7e2-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-260">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-261">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="ce7e2-261">Office for Mac</span></span></td>
    <td><span data-ttu-id="ce7e2-262">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-262">- Taskpane</span></span><br><span data-ttu-id="ce7e2-263">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="ce7e2-263">
        - Content</span></span><br><span data-ttu-id="ce7e2-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ce7e2-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ce7e2-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ce7e2-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ce7e2-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ce7e2-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ce7e2-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ce7e2-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-271">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="ce7e2-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ce7e2-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ce7e2-274">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-274">-BindingEvents</span></span><br><span data-ttu-id="ce7e2-275">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-275">
        -CompressedFile</span></span><br><span data-ttu-id="ce7e2-276">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-276">
        -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-277">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-277">
        - File</span></span><br><span data-ttu-id="ce7e2-278">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-278">
        -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-279">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-279">
        -MatrixBindings</span></span><br><span data-ttu-id="ce7e2-280">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-280">
        -MatrixCoercion</span></span><br><span data-ttu-id="ce7e2-281">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-281">
        -PdfFile</span></span><br><span data-ttu-id="ce7e2-282">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-282">
        - Selection</span></span><br><span data-ttu-id="ce7e2-283">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-283">
        - Settings</span></span><br><span data-ttu-id="ce7e2-284">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-284">
        -TableBindings</span></span><br><span data-ttu-id="ce7e2-285">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-285">
        -TableCoercion</span></span><br><span data-ttu-id="ce7e2-286">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-286">
        -TextBindings</span></span><br><span data-ttu-id="ce7e2-287">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-287">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="ce7e2-288">Outlook</span><span class="sxs-lookup"><span data-stu-id="ce7e2-288">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ce7e2-289">平台</span><span class="sxs-lookup"><span data-stu-id="ce7e2-289">Platform</span></span></th>
    <th><span data-ttu-id="ce7e2-290">扩展点</span><span class="sxs-lookup"><span data-stu-id="ce7e2-290">Extension points</span></span></th>
    <th><span data-ttu-id="ce7e2-291">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ce7e2-291">API requirement sets</span></span></th>
    <th><span data-ttu-id="ce7e2-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-293">Office Online</span><span class="sxs-lookup"><span data-stu-id="ce7e2-293">Office Online</span></span></td>
    <td> <span data-ttu-id="ce7e2-294">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ce7e2-294">- Mail Read</span></span><br><span data-ttu-id="ce7e2-295">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="ce7e2-295">
      - Mail Compose</span></span><br><span data-ttu-id="ce7e2-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ce7e2-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ce7e2-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ce7e2-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ce7e2-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ce7e2-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ce7e2-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="ce7e2-304">不适用</span><span class="sxs-lookup"><span data-stu-id="ce7e2-304">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-305">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="ce7e2-305">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="ce7e2-306">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ce7e2-306">- Mail Read</span></span><br><span data-ttu-id="ce7e2-307">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="ce7e2-307">
      - Mail Compose</span></span><br><span data-ttu-id="ce7e2-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ce7e2-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ce7e2-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ce7e2-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="ce7e2-313">不适用</span><span class="sxs-lookup"><span data-stu-id="ce7e2-313">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-314">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="ce7e2-314">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="ce7e2-315">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ce7e2-315">- Mail Read</span></span><br><span data-ttu-id="ce7e2-316">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="ce7e2-316">
      - Mail Compose</span></span><br><span data-ttu-id="ce7e2-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ce7e2-318">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="ce7e2-318">
      - Modules</span></span></td>
    <td> <span data-ttu-id="ce7e2-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ce7e2-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ce7e2-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ce7e2-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ce7e2-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ce7e2-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ce7e2-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="ce7e2-326">不适用</span><span class="sxs-lookup"><span data-stu-id="ce7e2-326">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-327">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="ce7e2-327">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="ce7e2-328">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ce7e2-328">- Mail Read</span></span><br><span data-ttu-id="ce7e2-329">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="ce7e2-329">
      - Mail Compose</span></span><br><span data-ttu-id="ce7e2-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ce7e2-331">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="ce7e2-331">
      - Modules</span></span></td>
    <td> <span data-ttu-id="ce7e2-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ce7e2-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ce7e2-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ce7e2-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ce7e2-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ce7e2-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ce7e2-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="ce7e2-339">不适用</span><span class="sxs-lookup"><span data-stu-id="ce7e2-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-340">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="ce7e2-340">Office for iOS</span></span></td>
    <td> <span data-ttu-id="ce7e2-341">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ce7e2-341">- Mail Read</span></span><br><span data-ttu-id="ce7e2-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ce7e2-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ce7e2-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ce7e2-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ce7e2-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ce7e2-348">不适用</span><span class="sxs-lookup"><span data-stu-id="ce7e2-348">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-349">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="ce7e2-349">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="ce7e2-350">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ce7e2-350">- Mail Read</span></span><br><span data-ttu-id="ce7e2-351">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="ce7e2-351">
      - Mail Compose</span></span><br><span data-ttu-id="ce7e2-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ce7e2-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ce7e2-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ce7e2-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ce7e2-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ce7e2-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ce7e2-359">不适用</span><span class="sxs-lookup"><span data-stu-id="ce7e2-359">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-360">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="ce7e2-360">Office for Mac</span></span></td>
    <td> <span data-ttu-id="ce7e2-361">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ce7e2-361">- Mail Read</span></span><br><span data-ttu-id="ce7e2-362">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="ce7e2-362">
      - Mail Compose</span></span><br><span data-ttu-id="ce7e2-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ce7e2-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ce7e2-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ce7e2-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ce7e2-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ce7e2-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ce7e2-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="ce7e2-371">不适用</span><span class="sxs-lookup"><span data-stu-id="ce7e2-371">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-372">Office for Android</span><span class="sxs-lookup"><span data-stu-id="ce7e2-372">Office for Android</span></span></td>
    <td> <span data-ttu-id="ce7e2-373">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="ce7e2-373">- Mail Read</span></span><br><span data-ttu-id="ce7e2-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-375">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-375">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ce7e2-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ce7e2-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ce7e2-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ce7e2-379">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-379">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ce7e2-380">不适用</span><span class="sxs-lookup"><span data-stu-id="ce7e2-380">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="ce7e2-381">Word</span><span class="sxs-lookup"><span data-stu-id="ce7e2-381">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ce7e2-382">平台</span><span class="sxs-lookup"><span data-stu-id="ce7e2-382">Platform</span></span></th>
    <th><span data-ttu-id="ce7e2-383">扩展点</span><span class="sxs-lookup"><span data-stu-id="ce7e2-383">Extension points</span></span></th>
    <th><span data-ttu-id="ce7e2-384">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ce7e2-384">API requirement sets</span></span></th>
    <th><span data-ttu-id="ce7e2-385"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-385"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-386">Office Online</span><span class="sxs-lookup"><span data-stu-id="ce7e2-386">Office Online</span></span></td>
    <td> <span data-ttu-id="ce7e2-387">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-387">- Taskpane</span></span><br><span data-ttu-id="ce7e2-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-389">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-389">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="ce7e2-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="ce7e2-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="ce7e2-392">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-392">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-393">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-393">-BindingEvents</span></span><br><span data-ttu-id="ce7e2-394">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ce7e2-394">
         -CustomXmlParts</span></span><br><span data-ttu-id="ce7e2-395">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-395">
         -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-396">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-396">
         - File</span></span><br><span data-ttu-id="ce7e2-397">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-397">
         -HtmlCoercion</span></span><br><span data-ttu-id="ce7e2-398">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-398">
         -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-399">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-399">
         -MatrixBindings</span></span><br><span data-ttu-id="ce7e2-400">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-400">
         -MatrixCoercion</span></span><br><span data-ttu-id="ce7e2-401">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-401">
         -OoxmlCoercion</span></span><br><span data-ttu-id="ce7e2-402">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-402">
         -PdfFile</span></span><br><span data-ttu-id="ce7e2-403">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-403">
         - Selection</span></span><br><span data-ttu-id="ce7e2-404">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-404">
         - Settings</span></span><br><span data-ttu-id="ce7e2-405">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-405">
         -TableBindings</span></span><br><span data-ttu-id="ce7e2-406">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-406">
         -TableCoercion</span></span><br><span data-ttu-id="ce7e2-407">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-407">
         -TextBindings</span></span><br><span data-ttu-id="ce7e2-408">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-408">
         -TextCoercion</span></span><br><span data-ttu-id="ce7e2-409">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-409">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-410">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="ce7e2-410">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="ce7e2-411">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-411">- Taskpane</span></span></td>
    <td> <span data-ttu-id="ce7e2-412">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-412">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-413">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-413">-BindingEvents</span></span><br><span data-ttu-id="ce7e2-414">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-414">
         -CompressedFile</span></span><br><span data-ttu-id="ce7e2-415">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ce7e2-415">
         -CustomXmlParts</span></span><br><span data-ttu-id="ce7e2-416">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-416">
         -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-417">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-417">
         - File</span></span><br><span data-ttu-id="ce7e2-418">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-418">
         -HtmlCoercion</span></span><br><span data-ttu-id="ce7e2-419">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-419">
         -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-420">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-420">
         -MatrixBindings</span></span><br><span data-ttu-id="ce7e2-421">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-421">
         -MatrixCoercion</span></span><br><span data-ttu-id="ce7e2-422">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-422">
         -OoxmlCoercion</span></span><br><span data-ttu-id="ce7e2-423">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-423">
         -PdfFile</span></span><br><span data-ttu-id="ce7e2-424">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-424">
         - Selection</span></span><br><span data-ttu-id="ce7e2-425">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-425">
         - Settings</span></span><br><span data-ttu-id="ce7e2-426">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-426">
         -TableBindings</span></span><br><span data-ttu-id="ce7e2-427">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-427">
         -TableCoercion</span></span><br><span data-ttu-id="ce7e2-428">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-428">
         -TextBindings</span></span><br><span data-ttu-id="ce7e2-429">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-429">
         -TextCoercion</span></span><br><span data-ttu-id="ce7e2-430">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-430">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-431">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="ce7e2-431">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="ce7e2-432">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-432">- Taskpane</span></span><br><span data-ttu-id="ce7e2-433">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-433">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-434">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-434">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="ce7e2-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="ce7e2-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="ce7e2-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-438">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-438">-BindingEvents</span></span><br><span data-ttu-id="ce7e2-439">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-439">
         -CompressedFile</span></span><br><span data-ttu-id="ce7e2-440">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ce7e2-440">
         -CustomXmlParts</span></span><br><span data-ttu-id="ce7e2-441">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-441">
         -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-442">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-442">
         - File</span></span><br><span data-ttu-id="ce7e2-443">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-443">
         -HtmlCoercion</span></span><br><span data-ttu-id="ce7e2-444">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-444">
         -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-445">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-445">
         -MatrixBindings</span></span><br><span data-ttu-id="ce7e2-446">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-446">
         -MatrixCoercion</span></span><br><span data-ttu-id="ce7e2-447">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-447">
         -OoxmlCoercion</span></span><br><span data-ttu-id="ce7e2-448">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-448">
         -PdfFile</span></span><br><span data-ttu-id="ce7e2-449">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-449">
         - Selection</span></span><br><span data-ttu-id="ce7e2-450">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-450">
         - Settings</span></span><br><span data-ttu-id="ce7e2-451">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-451">
         -TableBindings</span></span><br><span data-ttu-id="ce7e2-452">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-452">
         -TableCoercion</span></span><br><span data-ttu-id="ce7e2-453">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-453">
         -TextBindings</span></span><br><span data-ttu-id="ce7e2-454">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-454">
         -TextCoercion</span></span><br><span data-ttu-id="ce7e2-455">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-455">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-456">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="ce7e2-456">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="ce7e2-457">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-457">- Taskpane</span></span><br><span data-ttu-id="ce7e2-458">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-458">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-459">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-459">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="ce7e2-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="ce7e2-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="ce7e2-462">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-462">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-463">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-463">-BindingEvents</span></span><br><span data-ttu-id="ce7e2-464">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-464">
         -CompressedFile</span></span><br><span data-ttu-id="ce7e2-465">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ce7e2-465">
         -CustomXmlParts</span></span><br><span data-ttu-id="ce7e2-466">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-466">
         -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-467">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-467">
         - File</span></span><br><span data-ttu-id="ce7e2-468">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-468">
         -HtmlCoercion</span></span><br><span data-ttu-id="ce7e2-469">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-469">
         -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-470">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-470">
         -MatrixBindings</span></span><br><span data-ttu-id="ce7e2-471">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-471">
         -MatrixCoercion</span></span><br><span data-ttu-id="ce7e2-472">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-472">
         -OoxmlCoercion</span></span><br><span data-ttu-id="ce7e2-473">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-473">
         -PdfFile</span></span><br><span data-ttu-id="ce7e2-474">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-474">
         - Selection</span></span><br><span data-ttu-id="ce7e2-475">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-475">
         - Settings</span></span><br><span data-ttu-id="ce7e2-476">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-476">
         -TableBindings</span></span><br><span data-ttu-id="ce7e2-477">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-477">
         -TableCoercion</span></span><br><span data-ttu-id="ce7e2-478">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-478">
         -TextBindings</span></span><br><span data-ttu-id="ce7e2-479">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-479">
         -TextCoercion</span></span><br><span data-ttu-id="ce7e2-480">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-480">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-481">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="ce7e2-481">Office for iOS</span></span></td>
    <td> <span data-ttu-id="ce7e2-482">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-482">- Taskpane</span></span></td>
    <td> <span data-ttu-id="ce7e2-483">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-483">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="ce7e2-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="ce7e2-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="ce7e2-486">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="ce7e2-486">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="ce7e2-487">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-487">-BindingEvents</span></span><br><span data-ttu-id="ce7e2-488">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-488">
         -CompressedFile</span></span><br><span data-ttu-id="ce7e2-489">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ce7e2-489">
         -CustomXmlParts</span></span><br><span data-ttu-id="ce7e2-490">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-490">
         -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-491">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-491">
         - File</span></span><br><span data-ttu-id="ce7e2-492">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-492">
         -HtmlCoercion</span></span><br><span data-ttu-id="ce7e2-493">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-493">
         -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-494">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-494">
         -MatrixBindings</span></span><br><span data-ttu-id="ce7e2-495">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-495">
         -MatrixCoercion</span></span><br><span data-ttu-id="ce7e2-496">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-496">
         -OoxmlCoercion</span></span><br><span data-ttu-id="ce7e2-497">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-497">
         -PdfFile</span></span><br><span data-ttu-id="ce7e2-498">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-498">
         - Selection</span></span><br><span data-ttu-id="ce7e2-499">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-499">
         - Settings</span></span><br><span data-ttu-id="ce7e2-500">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-500">
         -TableBindings</span></span><br><span data-ttu-id="ce7e2-501">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-501">
         -TableCoercion</span></span><br><span data-ttu-id="ce7e2-502">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-502">
         -TextBindings</span></span><br><span data-ttu-id="ce7e2-503">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-503">
         -TextCoercion</span></span><br><span data-ttu-id="ce7e2-504">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-504">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-505">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="ce7e2-505">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="ce7e2-506">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-506">- Taskpane</span></span><br><span data-ttu-id="ce7e2-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="ce7e2-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="ce7e2-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="ce7e2-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="ce7e2-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="ce7e2-512">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-512">-BindingEvents</span></span><br><span data-ttu-id="ce7e2-513">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-513">
         -CompressedFile</span></span><br><span data-ttu-id="ce7e2-514">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ce7e2-514">
         -CustomXmlParts</span></span><br><span data-ttu-id="ce7e2-515">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-515">
         -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-516">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-516">
         - File</span></span><br><span data-ttu-id="ce7e2-517">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-517">
         -HtmlCoercion</span></span><br><span data-ttu-id="ce7e2-518">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-518">
         -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-519">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-519">
         -MatrixBindings</span></span><br><span data-ttu-id="ce7e2-520">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-520">
         -MatrixCoercion</span></span><br><span data-ttu-id="ce7e2-521">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-521">
         -OoxmlCoercion</span></span><br><span data-ttu-id="ce7e2-522">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-522">
         -PdfFile</span></span><br><span data-ttu-id="ce7e2-523">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-523">
         - Selection</span></span><br><span data-ttu-id="ce7e2-524">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-524">
         - Settings</span></span><br><span data-ttu-id="ce7e2-525">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-525">
         -TableBindings</span></span><br><span data-ttu-id="ce7e2-526">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-526">
         -TableCoercion</span></span><br><span data-ttu-id="ce7e2-527">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-527">
         -TextBindings</span></span><br><span data-ttu-id="ce7e2-528">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-528">
         -TextCoercion</span></span><br><span data-ttu-id="ce7e2-529">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-529">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-530">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="ce7e2-530">Office for Mac</span></span></td>
    <td> <span data-ttu-id="ce7e2-531">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-531">- Taskpane</span></span><br><span data-ttu-id="ce7e2-532">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-532">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-533">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-533">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="ce7e2-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="ce7e2-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="ce7e2-536">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="ce7e2-536">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="ce7e2-537">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-537">-BindingEvents</span></span><br><span data-ttu-id="ce7e2-538">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-538">
         -CompressedFile</span></span><br><span data-ttu-id="ce7e2-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ce7e2-539">
         -CustomXmlParts</span></span><br><span data-ttu-id="ce7e2-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-540">
         -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-541">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-541">
         - File</span></span><br><span data-ttu-id="ce7e2-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-542">
         -HtmlCoercion</span></span><br><span data-ttu-id="ce7e2-543">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-543">
         -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-544">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-544">
         -MatrixBindings</span></span><br><span data-ttu-id="ce7e2-545">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-545">
         -MatrixCoercion</span></span><br><span data-ttu-id="ce7e2-546">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-546">
         -OoxmlCoercion</span></span><br><span data-ttu-id="ce7e2-547">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-547">
         -PdfFile</span></span><br><span data-ttu-id="ce7e2-548">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-548">
         - Selection</span></span><br><span data-ttu-id="ce7e2-549">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-549">
         - Settings</span></span><br><span data-ttu-id="ce7e2-550">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-550">
         -TableBindings</span></span><br><span data-ttu-id="ce7e2-551">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-551">
         -TableCoercion</span></span><br><span data-ttu-id="ce7e2-552">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-552">
         -TextBindings</span></span><br><span data-ttu-id="ce7e2-553">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-553">
         -TextCoercion</span></span><br><span data-ttu-id="ce7e2-554">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-554">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="ce7e2-555">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="ce7e2-555">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ce7e2-556">平台</span><span class="sxs-lookup"><span data-stu-id="ce7e2-556">Platform</span></span></th>
    <th><span data-ttu-id="ce7e2-557">扩展点</span><span class="sxs-lookup"><span data-stu-id="ce7e2-557">Extension points</span></span></th>
    <th><span data-ttu-id="ce7e2-558">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ce7e2-558">API requirement sets</span></span></th>
    <th><span data-ttu-id="ce7e2-559"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-559"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-560">Office Online</span><span class="sxs-lookup"><span data-stu-id="ce7e2-560">Office Online</span></span></td>
    <td> <span data-ttu-id="ce7e2-561">- 内容</span><span class="sxs-lookup"><span data-stu-id="ce7e2-561">- Content</span></span><br><span data-ttu-id="ce7e2-562">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-562">
         - Taskpane</span></span><br><span data-ttu-id="ce7e2-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-565">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ce7e2-565">-ActiveView</span></span><br><span data-ttu-id="ce7e2-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-566">
         -CompressedFile</span></span><br><span data-ttu-id="ce7e2-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-567">
         -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-568">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-568">
         - File</span></span><br><span data-ttu-id="ce7e2-569">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-569">
         -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-570">
         -PdfFile</span></span><br><span data-ttu-id="ce7e2-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-571">
         - Selection</span></span><br><span data-ttu-id="ce7e2-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-572">
         - Settings</span></span><br><span data-ttu-id="ce7e2-573">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-573">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-574">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="ce7e2-574">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="ce7e2-575">- 内容</span><span class="sxs-lookup"><span data-stu-id="ce7e2-575">- Content</span></span><br><span data-ttu-id="ce7e2-576">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-576">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="ce7e2-577">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="ce7e2-577">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="ce7e2-578">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ce7e2-578">-ActiveView</span></span><br><span data-ttu-id="ce7e2-579">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-579">
         -CompressedFile</span></span><br><span data-ttu-id="ce7e2-580">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-580">
         -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-581">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-581">
         - File</span></span><br><span data-ttu-id="ce7e2-582">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-582">
         -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-583">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-583">
         -PdfFile</span></span><br><span data-ttu-id="ce7e2-584">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-584">
         - Selection</span></span><br><span data-ttu-id="ce7e2-585">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-585">
         - Settings</span></span><br><span data-ttu-id="ce7e2-586">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-586">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-587">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="ce7e2-587">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="ce7e2-588">- 内容</span><span class="sxs-lookup"><span data-stu-id="ce7e2-588">- Content</span></span><br><span data-ttu-id="ce7e2-589">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-589">
         - Taskpane</span></span><br><span data-ttu-id="ce7e2-590">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-590">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-592">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ce7e2-592">-ActiveView</span></span><br><span data-ttu-id="ce7e2-593">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-593">
         -CompressedFile</span></span><br><span data-ttu-id="ce7e2-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-594">
         -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-595">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-595">
         - File</span></span><br><span data-ttu-id="ce7e2-596">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-596">
         -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-597">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-597">
         -PdfFile</span></span><br><span data-ttu-id="ce7e2-598">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-598">
         - Selection</span></span><br><span data-ttu-id="ce7e2-599">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-599">
         - Settings</span></span><br><span data-ttu-id="ce7e2-600">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-600">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-601">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="ce7e2-601">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="ce7e2-602">- 内容</span><span class="sxs-lookup"><span data-stu-id="ce7e2-602">- Content</span></span><br><span data-ttu-id="ce7e2-603">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-603">
         - Taskpane</span></span><br><span data-ttu-id="ce7e2-604">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-604">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-605">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-605">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-606">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ce7e2-606">-ActiveView</span></span><br><span data-ttu-id="ce7e2-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-607">
         -CompressedFile</span></span><br><span data-ttu-id="ce7e2-608">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-608">
         -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-609">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-609">
         - File</span></span><br><span data-ttu-id="ce7e2-610">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-610">
         -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-611">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-611">
         -PdfFile</span></span><br><span data-ttu-id="ce7e2-612">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-612">
         - Selection</span></span><br><span data-ttu-id="ce7e2-613">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-613">
         - Settings</span></span><br><span data-ttu-id="ce7e2-614">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-614">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-615">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="ce7e2-615">Office for iOS</span></span></td>
    <td> <span data-ttu-id="ce7e2-616">- 内容</span><span class="sxs-lookup"><span data-stu-id="ce7e2-616">- Content</span></span><br><span data-ttu-id="ce7e2-617">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-617">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="ce7e2-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="ce7e2-619">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ce7e2-619">-ActiveView</span></span><br><span data-ttu-id="ce7e2-620">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-620">
         -CompressedFile</span></span><br><span data-ttu-id="ce7e2-621">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-621">
         -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-622">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-622">
         - File</span></span><br><span data-ttu-id="ce7e2-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-623">
         -PdfFile</span></span><br><span data-ttu-id="ce7e2-624">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-624">
         - Selection</span></span><br><span data-ttu-id="ce7e2-625">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-625">
         - Settings</span></span><br><span data-ttu-id="ce7e2-626">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-626">
         -TextCoercion</span></span><br><span data-ttu-id="ce7e2-627">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-627">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-628">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="ce7e2-628">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="ce7e2-629">- 内容</span><span class="sxs-lookup"><span data-stu-id="ce7e2-629">- Content</span></span><br><span data-ttu-id="ce7e2-630">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-630">
         - Taskpane</span></span><br><span data-ttu-id="ce7e2-631">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-631">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-632">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-632">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-633">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ce7e2-633">-ActiveView</span></span><br><span data-ttu-id="ce7e2-634">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-634">
         -CompressedFile</span></span><br><span data-ttu-id="ce7e2-635">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-635">
         -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-636">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-636">
         - File</span></span><br><span data-ttu-id="ce7e2-637">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-637">
         -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-638">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-638">
         -PdfFile</span></span><br><span data-ttu-id="ce7e2-639">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-639">
         - Selection</span></span><br><span data-ttu-id="ce7e2-640">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-640">
         - Settings</span></span><br><span data-ttu-id="ce7e2-641">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-641">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-642">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="ce7e2-642">Office for Mac</span></span></td>
    <td> <span data-ttu-id="ce7e2-643">- 内容</span><span class="sxs-lookup"><span data-stu-id="ce7e2-643">- Content</span></span><br><span data-ttu-id="ce7e2-644">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-644">
         - Taskpane</span></span><br><span data-ttu-id="ce7e2-645">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-645">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-646">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-646">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-647">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ce7e2-647">-ActiveView</span></span><br><span data-ttu-id="ce7e2-648">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-648">
         -CompressedFile</span></span><br><span data-ttu-id="ce7e2-649">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-649">
         -DocumentEvents</span></span><br><span data-ttu-id="ce7e2-650">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="ce7e2-650">
         - File</span></span><br><span data-ttu-id="ce7e2-651">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-651">
         -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-652">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ce7e2-652">
         -PdfFile</span></span><br><span data-ttu-id="ce7e2-653">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ce7e2-653">
         - Selection</span></span><br><span data-ttu-id="ce7e2-654">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-654">
         - Settings</span></span><br><span data-ttu-id="ce7e2-655">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-655">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="ce7e2-656">OneNote</span><span class="sxs-lookup"><span data-stu-id="ce7e2-656">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ce7e2-657">平台</span><span class="sxs-lookup"><span data-stu-id="ce7e2-657">Platform</span></span></th>
    <th><span data-ttu-id="ce7e2-658">扩展点</span><span class="sxs-lookup"><span data-stu-id="ce7e2-658">Extension points</span></span></th>
    <th><span data-ttu-id="ce7e2-659">API 要求集</span><span class="sxs-lookup"><span data-stu-id="ce7e2-659">API requirement sets</span></span></th>
    <th><span data-ttu-id="ce7e2-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="ce7e2-661">Office Online</span><span class="sxs-lookup"><span data-stu-id="ce7e2-661">Office Online</span></span></td>
    <td> <span data-ttu-id="ce7e2-662">- 内容</span><span class="sxs-lookup"><span data-stu-id="ce7e2-662">- Content</span></span><br><span data-ttu-id="ce7e2-663">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="ce7e2-663">
         - Taskpane</span></span><br><span data-ttu-id="ce7e2-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="ce7e2-666">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ce7e2-666">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ce7e2-667">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ce7e2-667">-DocumentEvents</span></span><br><span data-ttu-id="ce7e2-668">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-668">
         -HtmlCoercion</span></span><br><span data-ttu-id="ce7e2-669">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-669">
         -ImageCoercion</span></span><br><span data-ttu-id="ce7e2-670">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ce7e2-670">
         - Settings</span></span><br><span data-ttu-id="ce7e2-671">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ce7e2-671">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="ce7e2-672">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ce7e2-672">See also</span></span>

- [<span data-ttu-id="ce7e2-673">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="ce7e2-673">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="ce7e2-674">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="ce7e2-674">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="ce7e2-675">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="ce7e2-675">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="ce7e2-676">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="ce7e2-676">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)

---
title: Office 加载项主机和平台可用性
description: Excel、Word、Outlook、PowerPoint 和 OneNote 支持的要求集。
ms.date: 09/24/2018
ms.openlocfilehash: b06602e35ec906866ad16d667036a4cbaff2d89e
ms.sourcegitcommit: 8ce9a8d7f41d96879c39cc5527a3007dff25bee8
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/24/2018
ms.locfileid: "24985821"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="9a65d-103">Office 加载项主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="9a65d-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="9a65d-104">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。</span><span class="sxs-lookup"><span data-stu-id="9a65d-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="9a65d-105">下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="9a65d-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="9a65d-106">如果表格单元格内有星号 (\*)，表示我们正在完善它。</span><span class="sxs-lookup"><span data-stu-id="9a65d-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="9a65d-107">有关 Project 或 Access 要求集，请参阅 [Office 通用要求集](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="9a65d-107">For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="9a65d-p103">通过 MSI 安装的 Office 2016 的生成号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="9a65d-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="9a65d-110">Excel</span><span class="sxs-lookup"><span data-stu-id="9a65d-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="9a65d-111">平台</span><span class="sxs-lookup"><span data-stu-id="9a65d-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="9a65d-112">扩展点</span><span class="sxs-lookup"><span data-stu-id="9a65d-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="9a65d-113">API 要求集</span><span class="sxs-lookup"><span data-stu-id="9a65d-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="9a65d-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="9a65d-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="9a65d-115">Office Online</span></span></td>
    <td> <span data-ttu-id="9a65d-116">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-116">- Taskpane</span></span><br><span data-ttu-id="9a65d-117">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="9a65d-117">
        - Content</span></span><br><span data-ttu-id="9a65d-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="9a65d-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="9a65d-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9a65d-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9a65d-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9a65d-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9a65d-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9a65d-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9a65d-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="9a65d-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9a65d-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-127">
        -BindingEvents</span></span><br><span data-ttu-id="9a65d-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-128">
        -CompressedFile</span></span><br><span data-ttu-id="9a65d-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-129">
        -DocumentEvents</span></span><br><span data-ttu-id="9a65d-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="9a65d-130">
        - File</span></span><br><span data-ttu-id="9a65d-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-131">
        -MatrixBindings</span></span><br><span data-ttu-id="9a65d-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-132">
        -MatrixCoercion</span></span><br><span data-ttu-id="9a65d-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9a65d-133">
        - Selection</span></span><br><span data-ttu-id="9a65d-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-134">
        - Settings</span></span><br><span data-ttu-id="9a65d-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-135">
        -TableBindings</span></span><br><span data-ttu-id="9a65d-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-136">
        -TableCoercion</span></span><br><span data-ttu-id="9a65d-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-137">
        -TextBindings</span></span><br><span data-ttu-id="9a65d-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-138">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-139">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="9a65d-139">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="9a65d-140">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-140">
        - Taskpane</span></span><br><span data-ttu-id="9a65d-141">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="9a65d-141">
        - Content</span></span></td>
    <td>  <span data-ttu-id="9a65d-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9a65d-143">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-143">
        -BindingEvents</span></span><br><span data-ttu-id="9a65d-144">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-144">
        -CompressedFile</span></span><br><span data-ttu-id="9a65d-145">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-145">
        -DocumentEvents</span></span><br><span data-ttu-id="9a65d-146">
        - File</span><span class="sxs-lookup"><span data-stu-id="9a65d-146">
        - File</span></span><br><span data-ttu-id="9a65d-147">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-147">
        -ImageCoercion</span></span><br><span data-ttu-id="9a65d-148">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-148">
        -MatrixBindings</span></span><br><span data-ttu-id="9a65d-149">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-149">
        -MatrixCoercion</span></span><br><span data-ttu-id="9a65d-150">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9a65d-150">
        - Selection</span></span><br><span data-ttu-id="9a65d-151">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-151">
        - Settings</span></span><br><span data-ttu-id="9a65d-152">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-152">
        -TableBindings</span></span><br><span data-ttu-id="9a65d-153">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-153">
        -TableCoercion</span></span><br><span data-ttu-id="9a65d-154">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-154">
        -TextBindings</span></span><br><span data-ttu-id="9a65d-155">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-155">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-156">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="9a65d-156">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="9a65d-157">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-157">- Taskpane</span></span><br><span data-ttu-id="9a65d-158">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="9a65d-158">
        - Content</span></span><br><span data-ttu-id="9a65d-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9a65d-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9a65d-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9a65d-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9a65d-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9a65d-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9a65d-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9a65d-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-166">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="9a65d-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9a65d-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-168">-BindingEvents</span></span><br><span data-ttu-id="9a65d-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-169">
        -CompressedFile</span></span><br><span data-ttu-id="9a65d-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-170">
        -DocumentEvents</span></span><br><span data-ttu-id="9a65d-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="9a65d-171">
        - File</span></span><br><span data-ttu-id="9a65d-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-172">
        -ImageCoercion</span></span><br><span data-ttu-id="9a65d-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-173">
        -MatrixBindings</span></span><br><span data-ttu-id="9a65d-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="9a65d-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9a65d-175">
        - Selection</span></span><br><span data-ttu-id="9a65d-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-176">
        - Settings</span></span><br><span data-ttu-id="9a65d-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-177">
        -TableBindings</span></span><br><span data-ttu-id="9a65d-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-178">
        -TableCoercion</span></span><br><span data-ttu-id="9a65d-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-179">
        -TextBindings</span></span><br><span data-ttu-id="9a65d-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-181">Office for Windows</span><span class="sxs-lookup"><span data-stu-id="9a65d-181">Office for Windows</span></span></td>
    <td><span data-ttu-id="9a65d-182">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-182">- Taskpane</span></span><br><span data-ttu-id="9a65d-183">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="9a65d-183">
        - Content</span></span><br><span data-ttu-id="9a65d-184">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-184">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9a65d-185">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-185">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9a65d-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9a65d-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9a65d-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9a65d-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9a65d-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9a65d-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-191">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="9a65d-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9a65d-193">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-193">-BindingEvents</span></span><br><span data-ttu-id="9a65d-194">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-194">
        -CompressedFile</span></span><br><span data-ttu-id="9a65d-195">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-195">
        -DocumentEvents</span></span><br><span data-ttu-id="9a65d-196">
        - File</span><span class="sxs-lookup"><span data-stu-id="9a65d-196">
        - File</span></span><br><span data-ttu-id="9a65d-197">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-197">
        -ImageCoercion</span></span><br><span data-ttu-id="9a65d-198">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-198">
        -MatrixBindings</span></span><br><span data-ttu-id="9a65d-199">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-199">
        -MatrixCoercion</span></span><br><span data-ttu-id="9a65d-200">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9a65d-200">
        - Selection</span></span><br><span data-ttu-id="9a65d-201">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-201">
        - Settings</span></span><br><span data-ttu-id="9a65d-202">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-202">
        -TableBindings</span></span><br><span data-ttu-id="9a65d-203">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-203">
        -TableCoercion</span></span><br><span data-ttu-id="9a65d-204">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-204">
        -TextBindings</span></span><br><span data-ttu-id="9a65d-205">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-205">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-206">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="9a65d-206">Office for iOS</span></span></td>
    <td><span data-ttu-id="9a65d-207">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-207">- Taskpane</span></span><br><span data-ttu-id="9a65d-208">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="9a65d-208">
        - Content</span></span></td>
    <td><span data-ttu-id="9a65d-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9a65d-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9a65d-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9a65d-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9a65d-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9a65d-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9a65d-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-215">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="9a65d-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9a65d-217">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-217">-BindingEvents</span></span><br><span data-ttu-id="9a65d-218">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-218">
        -CompressedFile</span></span><br><span data-ttu-id="9a65d-219">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-219">
        -DocumentEvents</span></span><br><span data-ttu-id="9a65d-220">
        - File</span><span class="sxs-lookup"><span data-stu-id="9a65d-220">
        - File</span></span><br><span data-ttu-id="9a65d-221">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-221">
        -ImageCoercion</span></span><br><span data-ttu-id="9a65d-222">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-222">
        -MatrixBindings</span></span><br><span data-ttu-id="9a65d-223">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-223">
        -MatrixCoercion</span></span><br><span data-ttu-id="9a65d-224">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9a65d-224">
        - Selection</span></span><br><span data-ttu-id="9a65d-225">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-225">
        - Settings</span></span><br><span data-ttu-id="9a65d-226">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-226">
        -TableBindings</span></span><br><span data-ttu-id="9a65d-227">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-227">
        -TableCoercion</span></span><br><span data-ttu-id="9a65d-228">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-228">
        -TextBindings</span></span><br><span data-ttu-id="9a65d-229">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-229">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-230">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="9a65d-230">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="9a65d-231">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-231">- Taskpane</span></span><br><span data-ttu-id="9a65d-232">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="9a65d-232">
        - Content</span></span><br><span data-ttu-id="9a65d-233">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-233">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9a65d-234">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-234">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9a65d-235">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-235">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9a65d-236">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-236">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9a65d-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9a65d-238">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-238">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9a65d-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9a65d-240">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-240">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="9a65d-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9a65d-242">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-242">-BindingEvents</span></span><br><span data-ttu-id="9a65d-243">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-243">
        -CompressedFile</span></span><br><span data-ttu-id="9a65d-244">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-244">
        -DocumentEvents</span></span><br><span data-ttu-id="9a65d-245">
        - File</span><span class="sxs-lookup"><span data-stu-id="9a65d-245">
        - File</span></span><br><span data-ttu-id="9a65d-246">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-246">
        -ImageCoercion</span></span><br><span data-ttu-id="9a65d-247">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-247">
        -MatrixBindings</span></span><br><span data-ttu-id="9a65d-248">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-248">
        -MatrixCoercion</span></span><br><span data-ttu-id="9a65d-249">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-249">
        -PdfFile</span></span><br><span data-ttu-id="9a65d-250">
        - 选择</span><span class="sxs-lookup"><span data-stu-id="9a65d-250">
        - Selection</span></span><br><span data-ttu-id="9a65d-251">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-251">
        - Settings</span></span><br><span data-ttu-id="9a65d-252">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-252">
        -TableBindings</span></span><br><span data-ttu-id="9a65d-253">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-253">
        -TableCoercion</span></span><br><span data-ttu-id="9a65d-254">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-254">
        -TextBindings</span></span><br><span data-ttu-id="9a65d-255">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-255">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-256">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="9a65d-256">Office for Mac</span></span></td>
    <td><span data-ttu-id="9a65d-257">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-257">- Taskpane</span></span><br><span data-ttu-id="9a65d-258">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="9a65d-258">
        - Content</span></span><br><span data-ttu-id="9a65d-259">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-259">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9a65d-260">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-260">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9a65d-261">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-261">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9a65d-262">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-262">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9a65d-263">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-263">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9a65d-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9a65d-265">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-265">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9a65d-266">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-266">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="9a65d-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9a65d-268">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-268">-BindingEvents</span></span><br><span data-ttu-id="9a65d-269">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-269">
        -CompressedFile</span></span><br><span data-ttu-id="9a65d-270">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-270">
        -DocumentEvents</span></span><br><span data-ttu-id="9a65d-271">
        - File</span><span class="sxs-lookup"><span data-stu-id="9a65d-271">
        - File</span></span><br><span data-ttu-id="9a65d-272">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-272">
        -ImageCoercion</span></span><br><span data-ttu-id="9a65d-273">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-273">
        -MatrixBindings</span></span><br><span data-ttu-id="9a65d-274">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-274">
        -MatrixCoercion</span></span><br><span data-ttu-id="9a65d-275">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-275">
        -PdfFile</span></span><br><span data-ttu-id="9a65d-276">
        - 选择</span><span class="sxs-lookup"><span data-stu-id="9a65d-276">
        - Selection</span></span><br><span data-ttu-id="9a65d-277">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-277">
        - Settings</span></span><br><span data-ttu-id="9a65d-278">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-278">
        -TableBindings</span></span><br><span data-ttu-id="9a65d-279">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-279">
        -TableCoercion</span></span><br><span data-ttu-id="9a65d-280">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-280">
        -TextBindings</span></span><br><span data-ttu-id="9a65d-281">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-281">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="9a65d-282">Outlook</span><span class="sxs-lookup"><span data-stu-id="9a65d-282">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9a65d-283">平台</span><span class="sxs-lookup"><span data-stu-id="9a65d-283">Platform</span></span></th>
    <th><span data-ttu-id="9a65d-284">扩展点</span><span class="sxs-lookup"><span data-stu-id="9a65d-284">Extension points</span></span></th>
    <th><span data-ttu-id="9a65d-285">API 要求集</span><span class="sxs-lookup"><span data-stu-id="9a65d-285">API requirement sets</span></span></th>
    <th><span data-ttu-id="9a65d-286"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="9a65d-286"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-287">Office Online</span><span class="sxs-lookup"><span data-stu-id="9a65d-287">Office Online</span></span></td>
    <td> <span data-ttu-id="9a65d-288">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="9a65d-288">- Mail Read</span></span><br><span data-ttu-id="9a65d-289">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="9a65d-289">
      - Mail Compose</span></span><br><span data-ttu-id="9a65d-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9a65d-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9a65d-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9a65d-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9a65d-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9a65d-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9a65d-297">不适用</span><span class="sxs-lookup"><span data-stu-id="9a65d-297">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-298">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="9a65d-298">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9a65d-299">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="9a65d-299">- Mail Read</span></span><br><span data-ttu-id="9a65d-300">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="9a65d-300">
      - Mail Compose</span></span><br><span data-ttu-id="9a65d-301">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-301">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-302">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-302">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9a65d-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9a65d-304">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-304">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9a65d-305">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-305">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="9a65d-306">不适用</span><span class="sxs-lookup"><span data-stu-id="9a65d-306">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-307">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="9a65d-307">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9a65d-308">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="9a65d-308">- Mail Read</span></span><br><span data-ttu-id="9a65d-309">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="9a65d-309">
      - Mail Compose</span></span><br><span data-ttu-id="9a65d-310">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-310">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="9a65d-311">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="9a65d-311">
      - Modules</span></span></td>
    <td> <span data-ttu-id="9a65d-312">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-312">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9a65d-313">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-313">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9a65d-314">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-314">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9a65d-315">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-315">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9a65d-316">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-316">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9a65d-317">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-317">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9a65d-318">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-318">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="9a65d-319">不适用</span><span class="sxs-lookup"><span data-stu-id="9a65d-319">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-320">Office for Windows</span><span class="sxs-lookup"><span data-stu-id="9a65d-320">Office for Windows</span></span></td>
    <td> <span data-ttu-id="9a65d-321">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="9a65d-321">- Mail Read</span></span><br><span data-ttu-id="9a65d-322">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="9a65d-322">
      - Mail Compose</span></span><br><span data-ttu-id="9a65d-323">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-323">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="9a65d-324">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="9a65d-324">
      - Modules</span></span></td>
    <td> <span data-ttu-id="9a65d-325">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-325">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9a65d-326">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-326">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9a65d-327">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-327">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9a65d-328">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-328">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9a65d-329">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-329">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9a65d-330">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-330">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9a65d-331">不适用</span><span class="sxs-lookup"><span data-stu-id="9a65d-331">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-332">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="9a65d-332">Office for iOS</span></span></td>
    <td> <span data-ttu-id="9a65d-333">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="9a65d-333">- Mail Read</span></span><br><span data-ttu-id="9a65d-334">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-334">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-335">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-335">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9a65d-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9a65d-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9a65d-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9a65d-339">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-339">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="9a65d-340">不适用</span><span class="sxs-lookup"><span data-stu-id="9a65d-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-341">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="9a65d-341">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="9a65d-342">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="9a65d-342">- Mail Read</span></span><br><span data-ttu-id="9a65d-343">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="9a65d-343">
      - Mail Compose</span></span><br><span data-ttu-id="9a65d-344">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-344">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-345">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-345">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9a65d-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9a65d-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9a65d-348">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-348">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9a65d-349">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-349">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9a65d-350">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-350">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9a65d-351">不适用</span><span class="sxs-lookup"><span data-stu-id="9a65d-351">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-352">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="9a65d-352">Office for Mac</span></span></td>
    <td> <span data-ttu-id="9a65d-353">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="9a65d-353">- Mail Read</span></span><br><span data-ttu-id="9a65d-354">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="9a65d-354">
      - Mail Compose</span></span><br><span data-ttu-id="9a65d-355">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-355">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-356">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-356">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9a65d-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9a65d-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9a65d-359">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-359">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9a65d-360">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-360">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9a65d-361">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-361">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9a65d-362">不适用</span><span class="sxs-lookup"><span data-stu-id="9a65d-362">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-363">Office for Android</span><span class="sxs-lookup"><span data-stu-id="9a65d-363">Office for Android</span></span></td>
    <td> <span data-ttu-id="9a65d-364">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="9a65d-364">- Mail Read</span></span><br><span data-ttu-id="9a65d-365">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-365">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-366">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-366">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9a65d-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9a65d-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9a65d-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9a65d-370">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-370">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="9a65d-371">不适用</span><span class="sxs-lookup"><span data-stu-id="9a65d-371">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="9a65d-372">Word</span><span class="sxs-lookup"><span data-stu-id="9a65d-372">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9a65d-373">平台</span><span class="sxs-lookup"><span data-stu-id="9a65d-373">Platform</span></span></th>
    <th><span data-ttu-id="9a65d-374">扩展点</span><span class="sxs-lookup"><span data-stu-id="9a65d-374">Extension points</span></span></th>
    <th><span data-ttu-id="9a65d-375">API 要求集</span><span class="sxs-lookup"><span data-stu-id="9a65d-375">API requirement sets</span></span></th>
    <th><span data-ttu-id="9a65d-376"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="9a65d-376"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-377">Office Online</span><span class="sxs-lookup"><span data-stu-id="9a65d-377">Office Online</span></span></td>
    <td> <span data-ttu-id="9a65d-378">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-378">- Taskpane</span></span><br><span data-ttu-id="9a65d-379">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-379">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-380">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-380">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9a65d-381">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-381">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9a65d-382">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-382">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9a65d-383">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-383">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9a65d-384">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-384">-BindingEvents</span></span><br><span data-ttu-id="9a65d-385">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9a65d-385">
         -</span></span><br><span data-ttu-id="9a65d-386">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-386">
         -DocumentEvents</span></span><br><span data-ttu-id="9a65d-387">
         - File</span><span class="sxs-lookup"><span data-stu-id="9a65d-387">
         - File</span></span><br><span data-ttu-id="9a65d-388">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-388">
         -HtmlCoercion</span></span><br><span data-ttu-id="9a65d-389">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-389">
         -ImageCoercion</span></span><br><span data-ttu-id="9a65d-390">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-390">
         -MatrixBindings</span></span><br><span data-ttu-id="9a65d-391">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-391">
         -MatrixCoercion</span></span><br><span data-ttu-id="9a65d-392">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-392">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9a65d-393">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-393">
         -PdfFile</span></span><br><span data-ttu-id="9a65d-394">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9a65d-394">
         - Selection</span></span><br><span data-ttu-id="9a65d-395">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-395">
         - Settings</span></span><br><span data-ttu-id="9a65d-396">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-396">
         -TableBindings</span></span><br><span data-ttu-id="9a65d-397">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-397">
         -TableCoercion</span></span><br><span data-ttu-id="9a65d-398">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-398">
         -TextBindings</span></span><br><span data-ttu-id="9a65d-399">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-399">
         -TextCoercion</span></span><br><span data-ttu-id="9a65d-400">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-400">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-401">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="9a65d-401">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9a65d-402">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-402">- Taskpane</span></span></td>
    <td> <span data-ttu-id="9a65d-403">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-403">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9a65d-404">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-404">-BindingEvents</span></span><br><span data-ttu-id="9a65d-405">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-405">
         -CompressedFile</span></span><br><span data-ttu-id="9a65d-406">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9a65d-406">
         -</span></span><br><span data-ttu-id="9a65d-407">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-407">
         -DocumentEvents</span></span><br><span data-ttu-id="9a65d-408">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9a65d-408">
         - File</span></span><br><span data-ttu-id="9a65d-409">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-409">
         -HtmlCoercion</span></span><br><span data-ttu-id="9a65d-410">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-410">
         -ImageCoercion</span></span><br><span data-ttu-id="9a65d-411">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-411">
         -MatrixBindings</span></span><br><span data-ttu-id="9a65d-412">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-412">
         -MatrixCoercion</span></span><br><span data-ttu-id="9a65d-413">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-413">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9a65d-414">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-414">
         -PdfFile</span></span><br><span data-ttu-id="9a65d-415">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9a65d-415">
         - Selection</span></span><br><span data-ttu-id="9a65d-416">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-416">
         - Settings</span></span><br><span data-ttu-id="9a65d-417">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-417">
         -TableBindings</span></span><br><span data-ttu-id="9a65d-418">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-418">
         -TableCoercion</span></span><br><span data-ttu-id="9a65d-419">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-419">
         -TextBindings</span></span><br><span data-ttu-id="9a65d-420">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-420">
         -TextCoercion</span></span><br><span data-ttu-id="9a65d-421">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-421">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-422">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="9a65d-422">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9a65d-423">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-423">- Taskpane</span></span><br><span data-ttu-id="9a65d-424">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-424">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-425">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-425">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9a65d-426">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-426">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9a65d-427">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-427">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9a65d-428">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-428">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9a65d-429">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-429">-BindingEvents</span></span><br><span data-ttu-id="9a65d-430">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-430">
         -CompressedFile</span></span><br><span data-ttu-id="9a65d-431">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9a65d-431">
         -</span></span><br><span data-ttu-id="9a65d-432">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-432">
         -DocumentEvents</span></span><br><span data-ttu-id="9a65d-433">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9a65d-433">
         - File</span></span><br><span data-ttu-id="9a65d-434">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-434">
         -HtmlCoercion</span></span><br><span data-ttu-id="9a65d-435">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-435">
         -ImageCoercion</span></span><br><span data-ttu-id="9a65d-436">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-436">
         -MatrixBindings</span></span><br><span data-ttu-id="9a65d-437">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-437">
         -MatrixCoercion</span></span><br><span data-ttu-id="9a65d-438">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-438">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9a65d-439">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-439">
         -PdfFile</span></span><br><span data-ttu-id="9a65d-440">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9a65d-440">
         - Selection</span></span><br><span data-ttu-id="9a65d-441">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-441">
         - Settings</span></span><br><span data-ttu-id="9a65d-442">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-442">
         -TableBindings</span></span><br><span data-ttu-id="9a65d-443">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-443">
         -TableCoercion</span></span><br><span data-ttu-id="9a65d-444">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-444">
         -TextBindings</span></span><br><span data-ttu-id="9a65d-445">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-445">
         -TextCoercion</span></span><br><span data-ttu-id="9a65d-446">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-446">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-447">Office for Windows</span><span class="sxs-lookup"><span data-stu-id="9a65d-447">Office for Windows</span></span></td>
    <td> <span data-ttu-id="9a65d-448">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-448">- Taskpane</span></span><br><span data-ttu-id="9a65d-449">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-449">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-450">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-450">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9a65d-451">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-451">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9a65d-452">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-452">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9a65d-453">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-453">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9a65d-454">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-454">-BindingEvents</span></span><br><span data-ttu-id="9a65d-455">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-455">
         -CompressedFile</span></span><br><span data-ttu-id="9a65d-456">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9a65d-456">
         -</span></span><br><span data-ttu-id="9a65d-457">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-457">
         -DocumentEvents</span></span><br><span data-ttu-id="9a65d-458">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9a65d-458">
         - File</span></span><br><span data-ttu-id="9a65d-459">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-459">
         -HtmlCoercion</span></span><br><span data-ttu-id="9a65d-460">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-460">
         -ImageCoercion</span></span><br><span data-ttu-id="9a65d-461">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-461">
         -MatrixBindings</span></span><br><span data-ttu-id="9a65d-462">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-462">
         -MatrixCoercion</span></span><br><span data-ttu-id="9a65d-463">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-463">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9a65d-464">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-464">
         -PdfFile</span></span><br><span data-ttu-id="9a65d-465">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9a65d-465">
         - Selection</span></span><br><span data-ttu-id="9a65d-466">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="9a65d-466">
         - Settings</span></span><br><span data-ttu-id="9a65d-467">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-467">
         -TableBindings</span></span><br><span data-ttu-id="9a65d-468">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-468">
         -TableCoercion</span></span><br><span data-ttu-id="9a65d-469">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-469">
         -TextBindings</span></span><br><span data-ttu-id="9a65d-470">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-470">
         -TextCoercion</span></span><br><span data-ttu-id="9a65d-471">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-471">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-472">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="9a65d-472">Office for iOS</span></span></td>
    <td> <span data-ttu-id="9a65d-473">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-473">- Taskpane</span></span></td>
    <td> <span data-ttu-id="9a65d-474">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-474">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9a65d-475">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-475">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9a65d-476">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-476">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9a65d-477">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="9a65d-477">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="9a65d-478">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-478">-BindingEvents</span></span><br><span data-ttu-id="9a65d-479">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-479">
         -CompressedFile</span></span><br><span data-ttu-id="9a65d-480">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9a65d-480">
         -</span></span><br><span data-ttu-id="9a65d-481">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-481">
         -DocumentEvents</span></span><br><span data-ttu-id="9a65d-482">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9a65d-482">
         - File</span></span><br><span data-ttu-id="9a65d-483">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-483">
         -HtmlCoercion</span></span><br><span data-ttu-id="9a65d-484">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-484">
         -ImageCoercion</span></span><br><span data-ttu-id="9a65d-485">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-485">
         -MatrixBindings</span></span><br><span data-ttu-id="9a65d-486">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-486">
         -MatrixCoercion</span></span><br><span data-ttu-id="9a65d-487">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-487">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9a65d-488">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-488">
         -PdfFile</span></span><br><span data-ttu-id="9a65d-489">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9a65d-489">
         - Selection</span></span><br><span data-ttu-id="9a65d-490">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-490">
         - Settings</span></span><br><span data-ttu-id="9a65d-491">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-491">
         -TableBindings</span></span><br><span data-ttu-id="9a65d-492">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-492">
         -TableCoercion</span></span><br><span data-ttu-id="9a65d-493">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-493">
         -TextBindings</span></span><br><span data-ttu-id="9a65d-494">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-494">
         -TextCoercion</span></span><br><span data-ttu-id="9a65d-495">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-495">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-496">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="9a65d-496">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="9a65d-497">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-497">- Taskpane</span></span><br><span data-ttu-id="9a65d-498">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-498">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-499">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-499">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9a65d-500">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-500">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9a65d-501">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-501">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9a65d-502">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="9a65d-502">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="9a65d-503">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-503">-BindingEvents</span></span><br><span data-ttu-id="9a65d-504">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-504">
         -CompressedFile</span></span><br><span data-ttu-id="9a65d-505">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9a65d-505">
         -</span></span><br><span data-ttu-id="9a65d-506">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-506">
         -DocumentEvents</span></span><br><span data-ttu-id="9a65d-507">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9a65d-507">
         - File</span></span><br><span data-ttu-id="9a65d-508">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-508">
         -HtmlCoercion</span></span><br><span data-ttu-id="9a65d-509">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-509">
         -ImageCoercion</span></span><br><span data-ttu-id="9a65d-510">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-510">
         -MatrixBindings</span></span><br><span data-ttu-id="9a65d-511">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-511">
         -MatrixCoercion</span></span><br><span data-ttu-id="9a65d-512">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-512">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9a65d-513">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-513">
         -PdfFile</span></span><br><span data-ttu-id="9a65d-514">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="9a65d-514">
         - Selection</span></span><br><span data-ttu-id="9a65d-515">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-515">
         - Settings</span></span><br><span data-ttu-id="9a65d-516">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-516">
         -TableBindings</span></span><br><span data-ttu-id="9a65d-517">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-517">
         -TableCoercion</span></span><br><span data-ttu-id="9a65d-518">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-518">
         -TextBindings</span></span><br><span data-ttu-id="9a65d-519">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-519">
         -TextCoercion</span></span><br><span data-ttu-id="9a65d-520">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-520">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-521">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="9a65d-521">Office for Mac</span></span></td>
    <td> <span data-ttu-id="9a65d-522">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-522">- Taskpane</span></span><br><span data-ttu-id="9a65d-523">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-523">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-524">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-524">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9a65d-525">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-525">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9a65d-526">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-526">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9a65d-527">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="9a65d-527">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="9a65d-528">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-528">-BindingEvents</span></span><br><span data-ttu-id="9a65d-529">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-529">
         -CompressedFile</span></span><br><span data-ttu-id="9a65d-530">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9a65d-530">
         -</span></span><br><span data-ttu-id="9a65d-531">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-531">
         -DocumentEvents</span></span><br><span data-ttu-id="9a65d-532">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="9a65d-532">
         - File</span></span><br><span data-ttu-id="9a65d-533">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-533">
         -HtmlCoercion</span></span><br><span data-ttu-id="9a65d-534">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-534">
         -ImageCoercion</span></span><br><span data-ttu-id="9a65d-535">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-535">
         -MatrixBindings</span></span><br><span data-ttu-id="9a65d-536">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-536">
         -MatrixCoercion</span></span><br><span data-ttu-id="9a65d-537">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-537">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9a65d-538">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-538">
         -PdfFile</span></span><br><span data-ttu-id="9a65d-539">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9a65d-539">
         - Selection</span></span><br><span data-ttu-id="9a65d-540">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-540">
         - Settings</span></span><br><span data-ttu-id="9a65d-541">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-541">
         -TableBindings</span></span><br><span data-ttu-id="9a65d-542">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-542">
         -TableCoercion</span></span><br><span data-ttu-id="9a65d-543">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9a65d-543">
         -TextBindings</span></span><br><span data-ttu-id="9a65d-544">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-544">
         -TextCoercion</span></span><br><span data-ttu-id="9a65d-545">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-545">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="9a65d-546">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="9a65d-546">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9a65d-547">平台</span><span class="sxs-lookup"><span data-stu-id="9a65d-547">Platform</span></span></th>
    <th><span data-ttu-id="9a65d-548">扩展点</span><span class="sxs-lookup"><span data-stu-id="9a65d-548">Extension points</span></span></th>
    <th><span data-ttu-id="9a65d-549">API 要求集</span><span class="sxs-lookup"><span data-stu-id="9a65d-549">API requirement sets</span></span></th>
    <th><span data-ttu-id="9a65d-550"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="9a65d-550"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-551">Office Online</span><span class="sxs-lookup"><span data-stu-id="9a65d-551">Office Online</span></span></td>
    <td> <span data-ttu-id="9a65d-552">- 内容</span><span class="sxs-lookup"><span data-stu-id="9a65d-552">- Content</span></span><br><span data-ttu-id="9a65d-553">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-553">
         - Taskpane</span></span><br><span data-ttu-id="9a65d-554">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-554">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-555">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-555">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9a65d-556">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9a65d-556">-ActiveView</span></span><br><span data-ttu-id="9a65d-557">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-557">
         -CompressedFile</span></span><br><span data-ttu-id="9a65d-558">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-558">
         -DocumentEvents</span></span><br><span data-ttu-id="9a65d-559">
         - File</span><span class="sxs-lookup"><span data-stu-id="9a65d-559">
         - File</span></span><br><span data-ttu-id="9a65d-560">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-560">
         -ImageCoercion</span></span><br><span data-ttu-id="9a65d-561">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-561">
         -PdfFile</span></span><br><span data-ttu-id="9a65d-562">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9a65d-562">
         - Selection</span></span><br><span data-ttu-id="9a65d-563">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-563">
         - Settings</span></span><br><span data-ttu-id="9a65d-564">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-564">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-565">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="9a65d-565">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9a65d-566">- 内容</span><span class="sxs-lookup"><span data-stu-id="9a65d-566">- Content</span></span><br><span data-ttu-id="9a65d-567">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-567">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="9a65d-568">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="9a65d-568">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="9a65d-569">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9a65d-569">-ActiveView</span></span><br><span data-ttu-id="9a65d-570">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-570">
         -CompressedFile</span></span><br><span data-ttu-id="9a65d-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-571">
         -DocumentEvents</span></span><br><span data-ttu-id="9a65d-572">
         - File</span><span class="sxs-lookup"><span data-stu-id="9a65d-572">
         - File</span></span><br><span data-ttu-id="9a65d-573">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-573">
         -ImageCoercion</span></span><br><span data-ttu-id="9a65d-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-574">
         -PdfFile</span></span><br><span data-ttu-id="9a65d-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9a65d-575">
         - Selection</span></span><br><span data-ttu-id="9a65d-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-576">
         - Settings</span></span><br><span data-ttu-id="9a65d-577">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-577">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-578">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="9a65d-578">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9a65d-579">- 内容</span><span class="sxs-lookup"><span data-stu-id="9a65d-579">- Content</span></span><br><span data-ttu-id="9a65d-580">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-580">
         - Taskpane</span></span><br><span data-ttu-id="9a65d-581">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-581">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-582">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-582">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9a65d-583">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9a65d-583">-ActiveView</span></span><br><span data-ttu-id="9a65d-584">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-584">
         -CompressedFile</span></span><br><span data-ttu-id="9a65d-585">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-585">
         -DocumentEvents</span></span><br><span data-ttu-id="9a65d-586">
         - File</span><span class="sxs-lookup"><span data-stu-id="9a65d-586">
         - File</span></span><br><span data-ttu-id="9a65d-587">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-587">
         -ImageCoercion</span></span><br><span data-ttu-id="9a65d-588">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-588">
         -PdfFile</span></span><br><span data-ttu-id="9a65d-589">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9a65d-589">
         - Selection</span></span><br><span data-ttu-id="9a65d-590">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-590">
         - Settings</span></span><br><span data-ttu-id="9a65d-591">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-591">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-592">Office for Windows</span><span class="sxs-lookup"><span data-stu-id="9a65d-592">Office for Windows</span></span></td>
    <td> <span data-ttu-id="9a65d-593">- 内容</span><span class="sxs-lookup"><span data-stu-id="9a65d-593">- Content</span></span><br><span data-ttu-id="9a65d-594">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-594">
         - Taskpane</span></span><br><span data-ttu-id="9a65d-595">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-595">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-596">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-596">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9a65d-597">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9a65d-597">-ActiveView</span></span><br><span data-ttu-id="9a65d-598">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-598">
         -CompressedFile</span></span><br><span data-ttu-id="9a65d-599">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-599">
         -DocumentEvents</span></span><br><span data-ttu-id="9a65d-600">
         - File</span><span class="sxs-lookup"><span data-stu-id="9a65d-600">
         - File</span></span><br><span data-ttu-id="9a65d-601">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-601">
         -ImageCoercion</span></span><br><span data-ttu-id="9a65d-602">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-602">
         -PdfFile</span></span><br><span data-ttu-id="9a65d-603">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9a65d-603">
         - Selection</span></span><br><span data-ttu-id="9a65d-604">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-604">
         - Settings</span></span><br><span data-ttu-id="9a65d-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-605">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-606">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="9a65d-606">Office for iOS</span></span></td>
    <td> <span data-ttu-id="9a65d-607">- 内容</span><span class="sxs-lookup"><span data-stu-id="9a65d-607">- Content</span></span><br><span data-ttu-id="9a65d-608">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-608">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="9a65d-609">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-609">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="9a65d-610">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9a65d-610">-ActiveView</span></span><br><span data-ttu-id="9a65d-611">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-611">
         -CompressedFile</span></span><br><span data-ttu-id="9a65d-612">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-612">
         -DocumentEvents</span></span><br><span data-ttu-id="9a65d-613">
         - File</span><span class="sxs-lookup"><span data-stu-id="9a65d-613">
         - File</span></span><br><span data-ttu-id="9a65d-614">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-614">
         -PdfFile</span></span><br><span data-ttu-id="9a65d-615">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9a65d-615">
         - Selection</span></span><br><span data-ttu-id="9a65d-616">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-616">
         - Settings</span></span><br><span data-ttu-id="9a65d-617">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-617">
         -TextCoercion</span></span><br><span data-ttu-id="9a65d-618">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-618">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-619">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="9a65d-619">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="9a65d-620">- 内容</span><span class="sxs-lookup"><span data-stu-id="9a65d-620">- Content</span></span><br><span data-ttu-id="9a65d-621">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-621">
         - Taskpane</span></span><br><span data-ttu-id="9a65d-622">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-622">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-623">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-623">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9a65d-624">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9a65d-624">-ActiveView</span></span><br><span data-ttu-id="9a65d-625">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-625">
         -CompressedFile</span></span><br><span data-ttu-id="9a65d-626">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-626">
         -DocumentEvents</span></span><br><span data-ttu-id="9a65d-627">
         - File</span><span class="sxs-lookup"><span data-stu-id="9a65d-627">
         - File</span></span><br><span data-ttu-id="9a65d-628">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-628">
         -ImageCoercion</span></span><br><span data-ttu-id="9a65d-629">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-629">
         -PdfFile</span></span><br><span data-ttu-id="9a65d-630">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9a65d-630">
         - Selection</span></span><br><span data-ttu-id="9a65d-631">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-631">
         - Settings</span></span><br><span data-ttu-id="9a65d-632">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-632">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-633">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="9a65d-633">Office for Mac</span></span></td>
    <td> <span data-ttu-id="9a65d-634">- 内容</span><span class="sxs-lookup"><span data-stu-id="9a65d-634">- Content</span></span><br><span data-ttu-id="9a65d-635">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-635">
         - Taskpane</span></span><br><span data-ttu-id="9a65d-636">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-636">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-637">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-637">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9a65d-638">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9a65d-638">-ActiveView</span></span><br><span data-ttu-id="9a65d-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-639">
         -CompressedFile</span></span><br><span data-ttu-id="9a65d-640">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-640">
         -DocumentEvents</span></span><br><span data-ttu-id="9a65d-641">
         - File</span><span class="sxs-lookup"><span data-stu-id="9a65d-641">
         - File</span></span><br><span data-ttu-id="9a65d-642">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-642">
         -ImageCoercion</span></span><br><span data-ttu-id="9a65d-643">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9a65d-643">
         -PdfFile</span></span><br><span data-ttu-id="9a65d-644">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9a65d-644">
         - Selection</span></span><br><span data-ttu-id="9a65d-645">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-645">
         - Settings</span></span><br><span data-ttu-id="9a65d-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-646">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="9a65d-647">OneNote</span><span class="sxs-lookup"><span data-stu-id="9a65d-647">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9a65d-648">平台</span><span class="sxs-lookup"><span data-stu-id="9a65d-648">Platform</span></span></th>
    <th><span data-ttu-id="9a65d-649">扩展点</span><span class="sxs-lookup"><span data-stu-id="9a65d-649">Extension points</span></span></th>
    <th><span data-ttu-id="9a65d-650">API 要求集</span><span class="sxs-lookup"><span data-stu-id="9a65d-650">API requirement sets</span></span></th>
    <th><span data-ttu-id="9a65d-651"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="9a65d-651"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="9a65d-652">Office Online</span><span class="sxs-lookup"><span data-stu-id="9a65d-652">Office Online</span></span></td>
    <td> <span data-ttu-id="9a65d-653">- 内容</span><span class="sxs-lookup"><span data-stu-id="9a65d-653">- Content</span></span><br><span data-ttu-id="9a65d-654">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="9a65d-654">
         - Taskpane</span></span><br><span data-ttu-id="9a65d-655">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-655">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9a65d-656">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-656">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="9a65d-657">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9a65d-657">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9a65d-658">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9a65d-658">-DocumentEvents</span></span><br><span data-ttu-id="9a65d-659">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-659">
         -HtmlCoercion</span></span><br><span data-ttu-id="9a65d-660">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-660">
         -ImageCoercion</span></span><br><span data-ttu-id="9a65d-661">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9a65d-661">
         - Settings</span></span><br><span data-ttu-id="9a65d-662">
         - contTextCoercion</span><span class="sxs-lookup"><span data-stu-id="9a65d-662">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="9a65d-663">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9a65d-663">See also</span></span>

- [<span data-ttu-id="9a65d-664">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="9a65d-664">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="9a65d-665">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="9a65d-665">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="9a65d-666">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="9a65d-666">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="9a65d-667">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="9a65d-667">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)

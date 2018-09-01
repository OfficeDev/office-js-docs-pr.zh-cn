---
title: Office 加载项主机和平台可用性
description: Excel、Word、Outlook、PowerPoint 和 OneNote 支持的要求集。
ms.date: 08/30/2018
ms.openlocfilehash: 06fb073693bd8adca7d196f4361699ac3f54cee1
ms.sourcegitcommit: 78b28ae88d53bfef3134c09cc4336a5a8722c70b
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/01/2018
ms.locfileid: "23797299"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="5e377-103">Office 加载项主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="5e377-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="5e377-104">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。</span><span class="sxs-lookup"><span data-stu-id="5e377-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="5e377-105">下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="5e377-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="5e377-106">如果表格单元格内有星号 (\*)，表示我们正在完善它。</span><span class="sxs-lookup"><span data-stu-id="5e377-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="5e377-107">有关 Project 或 Access 要求集，请参阅 [Office 通用要求集](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="5e377-107">For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="5e377-p103">通过 MSI 安装的 Office 2016 的生成号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="5e377-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="5e377-110">Excel</span><span class="sxs-lookup"><span data-stu-id="5e377-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="5e377-111">平台</span><span class="sxs-lookup"><span data-stu-id="5e377-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="5e377-112">扩展点</span><span class="sxs-lookup"><span data-stu-id="5e377-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="5e377-113">API 要求集</span><span class="sxs-lookup"><span data-stu-id="5e377-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="5e377-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5e377-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="5e377-115">Office Online</span></span></td>
    <td> <span data-ttu-id="5e377-116">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-116">- Taskpane</span></span><br><span data-ttu-id="5e377-117">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="5e377-117">
        - Content</span></span><br><span data-ttu-id="5e377-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="5e377-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="5e377-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5e377-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5e377-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5e377-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5e377-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5e377-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5e377-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5e377-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5e377-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5e377-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5e377-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5e377-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5e377-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="5e377-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5e377-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-127">
        -BindingEvents</span></span><br><span data-ttu-id="5e377-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5e377-128">
        -CompressedFile</span></span><br><span data-ttu-id="5e377-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-129">
        -DocumentEvents</span></span><br><span data-ttu-id="5e377-130">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="5e377-130">
        - File</span></span><br><span data-ttu-id="5e377-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-131">
        -MatrixBindings</span></span><br><span data-ttu-id="5e377-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-132">
        -MatrixCoercion</span></span><br><span data-ttu-id="5e377-133">
        - 选择</span><span class="sxs-lookup"><span data-stu-id="5e377-133">
        - Selection</span></span><br><span data-ttu-id="5e377-134">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-134">
        - Settings</span></span><br><span data-ttu-id="5e377-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-135">
        -TableBindings</span></span><br><span data-ttu-id="5e377-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-136">
        -TableCoercion</span></span><br><span data-ttu-id="5e377-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-137">
        -TextBindings</span></span><br><span data-ttu-id="5e377-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-138">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-139">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="5e377-139">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="5e377-140">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-140">
        - Taskpane</span></span><br><span data-ttu-id="5e377-141">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="5e377-141">
        - Content</span></span></td>
    <td>  <span data-ttu-id="5e377-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5e377-143">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-143">
        -BindingEvents</span></span><br><span data-ttu-id="5e377-144">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5e377-144">
        -CompressedFile</span></span><br><span data-ttu-id="5e377-145">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-145">
        -DocumentEvents</span></span><br><span data-ttu-id="5e377-146">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="5e377-146">
        - File</span></span><br><span data-ttu-id="5e377-147">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-147">
        -ImageCoercion</span></span><br><span data-ttu-id="5e377-148">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-148">
        -MatrixBindings</span></span><br><span data-ttu-id="5e377-149">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-149">
        -MatrixCoercion</span></span><br><span data-ttu-id="5e377-150">
        - 选择</span><span class="sxs-lookup"><span data-stu-id="5e377-150">
        - Selection</span></span><br><span data-ttu-id="5e377-151">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-151">
        - Settings</span></span><br><span data-ttu-id="5e377-152">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-152">
        -TableBindings</span></span><br><span data-ttu-id="5e377-153">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-153">
        -TableCoercion</span></span><br><span data-ttu-id="5e377-154">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-154">
        -TextBindings</span></span><br><span data-ttu-id="5e377-155">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-155">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-156">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="5e377-156">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="5e377-157">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-157">- Taskpane</span></span><br><span data-ttu-id="5e377-158">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="5e377-158">
        - Content</span></span><br><span data-ttu-id="5e377-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="5e377-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5e377-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5e377-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5e377-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5e377-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5e377-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5e377-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5e377-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5e377-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5e377-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5e377-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5e377-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5e377-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5e377-166">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="5e377-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5e377-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-168">-BindingEvents</span></span><br><span data-ttu-id="5e377-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5e377-169">
        -CompressedFile</span></span><br><span data-ttu-id="5e377-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-170">
        -DocumentEvents</span></span><br><span data-ttu-id="5e377-171">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="5e377-171">
        - File</span></span><br><span data-ttu-id="5e377-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-172">
        -ImageCoercion</span></span><br><span data-ttu-id="5e377-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-173">
        -MatrixBindings</span></span><br><span data-ttu-id="5e377-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="5e377-175">
        - 选择</span><span class="sxs-lookup"><span data-stu-id="5e377-175">
        - Selection</span></span><br><span data-ttu-id="5e377-176">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-176">
        - Settings</span></span><br><span data-ttu-id="5e377-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-177">
        -TableBindings</span></span><br><span data-ttu-id="5e377-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-178">
        -TableCoercion</span></span><br><span data-ttu-id="5e377-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-179">
        -TextBindings</span></span><br><span data-ttu-id="5e377-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-181">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="5e377-181">Office for iOS</span></span></td>
    <td><span data-ttu-id="5e377-182">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-182">- Taskpane</span></span><br><span data-ttu-id="5e377-183">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="5e377-183">
        - Content</span></span></td>
    <td><span data-ttu-id="5e377-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5e377-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5e377-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5e377-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5e377-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5e377-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5e377-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5e377-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5e377-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5e377-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5e377-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5e377-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5e377-190">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="5e377-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5e377-192">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-192">-BindingEvents</span></span><br><span data-ttu-id="5e377-193">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5e377-193">
        -CompressedFile</span></span><br><span data-ttu-id="5e377-194">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-194">
        -DocumentEvents</span></span><br><span data-ttu-id="5e377-195">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="5e377-195">
        - File</span></span><br><span data-ttu-id="5e377-196">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-196">
        -ImageCoercion</span></span><br><span data-ttu-id="5e377-197">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-197">
        -MatrixBindings</span></span><br><span data-ttu-id="5e377-198">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-198">
        -MatrixCoercion</span></span><br><span data-ttu-id="5e377-199">
        - 选择</span><span class="sxs-lookup"><span data-stu-id="5e377-199">
        - Selection</span></span><br><span data-ttu-id="5e377-200">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-200">
        - Settings</span></span><br><span data-ttu-id="5e377-201">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-201">
        -TableBindings</span></span><br><span data-ttu-id="5e377-202">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-202">
        -TableCoercion</span></span><br><span data-ttu-id="5e377-203">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-203">
        -TextBindings</span></span><br><span data-ttu-id="5e377-204">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-204">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-205">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="5e377-205">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="5e377-206">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-206">- Taskpane</span></span><br><span data-ttu-id="5e377-207">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="5e377-207">
        - Content</span></span><br><span data-ttu-id="5e377-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="5e377-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5e377-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5e377-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5e377-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5e377-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5e377-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5e377-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5e377-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5e377-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5e377-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5e377-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5e377-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5e377-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5e377-215">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="5e377-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5e377-217">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-217">-BindingEvents</span></span><br><span data-ttu-id="5e377-218">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5e377-218">
        -CompressedFile</span></span><br><span data-ttu-id="5e377-219">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-219">
        -DocumentEvents</span></span><br><span data-ttu-id="5e377-220">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="5e377-220">
        - File</span></span><br><span data-ttu-id="5e377-221">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-221">
        -ImageCoercion</span></span><br><span data-ttu-id="5e377-222">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-222">
        -MatrixBindings</span></span><br><span data-ttu-id="5e377-223">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-223">
        -MatrixCoercion</span></span><br><span data-ttu-id="5e377-224">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5e377-224">
        -PdfFile</span></span><br><span data-ttu-id="5e377-225">
        - 选择</span><span class="sxs-lookup"><span data-stu-id="5e377-225">
        - Selection</span></span><br><span data-ttu-id="5e377-226">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-226">
        - Settings</span></span><br><span data-ttu-id="5e377-227">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-227">
        -TableBindings</span></span><br><span data-ttu-id="5e377-228">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-228">
        -TableCoercion</span></span><br><span data-ttu-id="5e377-229">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-229">
        -TextBindings</span></span><br><span data-ttu-id="5e377-230">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-230">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="5e377-231">Outlook</span><span class="sxs-lookup"><span data-stu-id="5e377-231">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5e377-232">平台</span><span class="sxs-lookup"><span data-stu-id="5e377-232">Platform</span></span></th>
    <th><span data-ttu-id="5e377-233">扩展点</span><span class="sxs-lookup"><span data-stu-id="5e377-233">Extension points</span></span></th>
    <th><span data-ttu-id="5e377-234">API 要求集</span><span class="sxs-lookup"><span data-stu-id="5e377-234">API requirement sets</span></span></th>
    <th><span data-ttu-id="5e377-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5e377-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-236">Office Online</span><span class="sxs-lookup"><span data-stu-id="5e377-236">Office Online</span></span></td>
    <td> <span data-ttu-id="5e377-237">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="5e377-237">- Mail Read</span></span><br><span data-ttu-id="5e377-238">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="5e377-238">
      - Mail Compose</span></span><br><span data-ttu-id="5e377-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="5e377-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5e377-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5e377-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5e377-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5e377-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5e377-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5e377-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5e377-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5e377-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5e377-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5e377-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5e377-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5e377-246">不适用</span><span class="sxs-lookup"><span data-stu-id="5e377-246">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-247">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="5e377-247">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="5e377-248">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="5e377-248">- Mail Read</span></span><br><span data-ttu-id="5e377-249">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="5e377-249">
      - Mail Compose</span></span><br><span data-ttu-id="5e377-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="5e377-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5e377-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5e377-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5e377-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5e377-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5e377-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5e377-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5e377-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="5e377-255">不适用</span><span class="sxs-lookup"><span data-stu-id="5e377-255">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-256">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="5e377-256">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="5e377-257">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="5e377-257">- Mail Read</span></span><br><span data-ttu-id="5e377-258">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="5e377-258">
      - Mail Compose</span></span><br><span data-ttu-id="5e377-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="5e377-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5e377-260">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="5e377-260">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5e377-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5e377-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5e377-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5e377-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5e377-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5e377-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5e377-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5e377-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5e377-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5e377-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5e377-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5e377-267">不适用</span><span class="sxs-lookup"><span data-stu-id="5e377-267">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-268">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="5e377-268">Office for iOS</span></span></td>
    <td> <span data-ttu-id="5e377-269">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="5e377-269">- Mail Read</span></span><br><span data-ttu-id="5e377-270">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="5e377-270">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5e377-271">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-271">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5e377-272">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5e377-272">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5e377-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5e377-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5e377-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5e377-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5e377-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5e377-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="5e377-276">不适用</span><span class="sxs-lookup"><span data-stu-id="5e377-276">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-277">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="5e377-277">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="5e377-278">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="5e377-278">- Mail Read</span></span><br><span data-ttu-id="5e377-279">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="5e377-279">
      - Mail Compose</span></span><br><span data-ttu-id="5e377-280">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="5e377-280">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5e377-281">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-281">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5e377-282">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5e377-282">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5e377-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5e377-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5e377-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5e377-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5e377-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5e377-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5e377-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5e377-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5e377-287">不适用</span><span class="sxs-lookup"><span data-stu-id="5e377-287">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-288">Office for Android</span><span class="sxs-lookup"><span data-stu-id="5e377-288">Office for Android</span></span></td>
    <td> <span data-ttu-id="5e377-289">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="5e377-289">- Mail Read</span></span><br><span data-ttu-id="5e377-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="5e377-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5e377-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5e377-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5e377-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5e377-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5e377-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5e377-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5e377-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5e377-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5e377-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="5e377-296">不适用</span><span class="sxs-lookup"><span data-stu-id="5e377-296">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="5e377-297">Word</span><span class="sxs-lookup"><span data-stu-id="5e377-297">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5e377-298">平台</span><span class="sxs-lookup"><span data-stu-id="5e377-298">Platform</span></span></th>
    <th><span data-ttu-id="5e377-299">扩展点</span><span class="sxs-lookup"><span data-stu-id="5e377-299">Extension points</span></span></th>
    <th><span data-ttu-id="5e377-300">API 要求集</span><span class="sxs-lookup"><span data-stu-id="5e377-300">API requirement sets</span></span></th>
    <th><span data-ttu-id="5e377-301"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5e377-301"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-302">Office Online</span><span class="sxs-lookup"><span data-stu-id="5e377-302">Office Online</span></span></td>
    <td> <span data-ttu-id="5e377-303">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-303">- Taskpane</span></span><br><span data-ttu-id="5e377-304">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="5e377-304">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5e377-305">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-305">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5e377-306">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5e377-306">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5e377-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5e377-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5e377-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5e377-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-309">-BindingEvents</span></span><br><span data-ttu-id="5e377-310">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5e377-310">
         -</span></span><br><span data-ttu-id="5e377-311">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-311">
         -DocumentEvents</span></span><br><span data-ttu-id="5e377-312">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="5e377-312">
         - File</span></span><br><span data-ttu-id="5e377-313">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-313">
         -HtmlCoercion</span></span><br><span data-ttu-id="5e377-314">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-314">
         -ImageCoercion</span></span><br><span data-ttu-id="5e377-315">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-315">
         -MatrixBindings</span></span><br><span data-ttu-id="5e377-316">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-316">
         -MatrixCoercion</span></span><br><span data-ttu-id="5e377-317">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-317">
         -OoxmlCoercion</span></span><br><span data-ttu-id="5e377-318">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5e377-318">
         -PdfFile</span></span><br><span data-ttu-id="5e377-319">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="5e377-319">
         - Selection</span></span><br><span data-ttu-id="5e377-320">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-320">
         - Settings</span></span><br><span data-ttu-id="5e377-321">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-321">
         -TableBindings</span></span><br><span data-ttu-id="5e377-322">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-322">
         -TableCoercion</span></span><br><span data-ttu-id="5e377-323">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-323">
         -TextBindings</span></span><br><span data-ttu-id="5e377-324">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-324">
         -TextCoercion</span></span><br><span data-ttu-id="5e377-325">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5e377-325">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-326">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="5e377-326">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="5e377-327">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-327">- Taskpane</span></span></td>
    <td> <span data-ttu-id="5e377-328">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-328">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5e377-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-329">-BindingEvents</span></span><br><span data-ttu-id="5e377-330">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5e377-330">
         -CompressedFile</span></span><br><span data-ttu-id="5e377-331">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5e377-331">
         -</span></span><br><span data-ttu-id="5e377-332">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-332">
         -DocumentEvents</span></span><br><span data-ttu-id="5e377-333">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="5e377-333">
         - File</span></span><br><span data-ttu-id="5e377-334">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-334">
         -HtmlCoercion</span></span><br><span data-ttu-id="5e377-335">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-335">
         -ImageCoercion</span></span><br><span data-ttu-id="5e377-336">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-336">
         -MatrixBindings</span></span><br><span data-ttu-id="5e377-337">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-337">
         -MatrixCoercion</span></span><br><span data-ttu-id="5e377-338">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-338">
         -OoxmlCoercion</span></span><br><span data-ttu-id="5e377-339">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5e377-339">
         -PdfFile</span></span><br><span data-ttu-id="5e377-340">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="5e377-340">
         - Selection</span></span><br><span data-ttu-id="5e377-341">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-341">
         - Settings</span></span><br><span data-ttu-id="5e377-342">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-342">
         -TableBindings</span></span><br><span data-ttu-id="5e377-343">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-343">
         -TableCoercion</span></span><br><span data-ttu-id="5e377-344">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-344">
         -TextBindings</span></span><br><span data-ttu-id="5e377-345">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-345">
         -TextCoercion</span></span><br><span data-ttu-id="5e377-346">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5e377-346">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-347">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="5e377-347">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="5e377-348">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-348">- Taskpane</span></span><br><span data-ttu-id="5e377-349">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="5e377-349">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5e377-350">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-350">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5e377-351">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5e377-351">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5e377-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5e377-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5e377-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5e377-354">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-354">-BindingEvents</span></span><br><span data-ttu-id="5e377-355">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5e377-355">
         -CompressedFile</span></span><br><span data-ttu-id="5e377-356">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5e377-356">
         -</span></span><br><span data-ttu-id="5e377-357">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-357">
         -DocumentEvents</span></span><br><span data-ttu-id="5e377-358">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="5e377-358">
         - File</span></span><br><span data-ttu-id="5e377-359">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-359">
         -HtmlCoercion</span></span><br><span data-ttu-id="5e377-360">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-360">
         -ImageCoercion</span></span><br><span data-ttu-id="5e377-361">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-361">
         -MatrixBindings</span></span><br><span data-ttu-id="5e377-362">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-362">
         -MatrixCoercion</span></span><br><span data-ttu-id="5e377-363">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-363">
         -OoxmlCoercion</span></span><br><span data-ttu-id="5e377-364">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5e377-364">
         -PdfFile</span></span><br><span data-ttu-id="5e377-365">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="5e377-365">
         - Selection</span></span><br><span data-ttu-id="5e377-366">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-366">
         - Settings</span></span><br><span data-ttu-id="5e377-367">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-367">
         -TableBindings</span></span><br><span data-ttu-id="5e377-368">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-368">
         -TableCoercion</span></span><br><span data-ttu-id="5e377-369">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-369">
         -TextBindings</span></span><br><span data-ttu-id="5e377-370">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-370">
         -TextCoercion</span></span><br><span data-ttu-id="5e377-371">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5e377-371">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-372">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="5e377-372">Office for iOS</span></span></td>
    <td> <span data-ttu-id="5e377-373">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-373">- Taskpane</span></span></td>
    <td> <span data-ttu-id="5e377-374">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-374">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5e377-375">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5e377-375">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5e377-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5e377-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5e377-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5e377-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5e377-378">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-378">-BindingEvents</span></span><br><span data-ttu-id="5e377-379">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5e377-379">
         -CompressedFile</span></span><br><span data-ttu-id="5e377-380">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5e377-380">
         -</span></span><br><span data-ttu-id="5e377-381">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-381">
         -DocumentEvents</span></span><br><span data-ttu-id="5e377-382">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="5e377-382">
         - File</span></span><br><span data-ttu-id="5e377-383">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-383">
         -HtmlCoercion</span></span><br><span data-ttu-id="5e377-384">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-384">
         -ImageCoercion</span></span><br><span data-ttu-id="5e377-385">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-385">
         -MatrixBindings</span></span><br><span data-ttu-id="5e377-386">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-386">
         -MatrixCoercion</span></span><br><span data-ttu-id="5e377-387">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-387">
         -OoxmlCoercion</span></span><br><span data-ttu-id="5e377-388">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5e377-388">
         -PdfFile</span></span><br><span data-ttu-id="5e377-389">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="5e377-389">
         - Selection</span></span><br><span data-ttu-id="5e377-390">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-390">
         - Settings</span></span><br><span data-ttu-id="5e377-391">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-391">
         -TableBindings</span></span><br><span data-ttu-id="5e377-392">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-392">
         -TableCoercion</span></span><br><span data-ttu-id="5e377-393">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-393">
         -TextBindings</span></span><br><span data-ttu-id="5e377-394">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-394">
         -TextCoercion</span></span><br><span data-ttu-id="5e377-395">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5e377-395">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-396">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="5e377-396">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="5e377-397">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-397">- Taskpane</span></span><br><span data-ttu-id="5e377-398">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="5e377-398">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5e377-399">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-399">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5e377-400">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5e377-400">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5e377-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5e377-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5e377-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5e377-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5e377-403">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-403">-BindingEvents</span></span><br><span data-ttu-id="5e377-404">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5e377-404">
         -CompressedFile</span></span><br><span data-ttu-id="5e377-405">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5e377-405">
         -</span></span><br><span data-ttu-id="5e377-406">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-406">
         -DocumentEvents</span></span><br><span data-ttu-id="5e377-407">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="5e377-407">
         - File</span></span><br><span data-ttu-id="5e377-408">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-408">
         -HtmlCoercion</span></span><br><span data-ttu-id="5e377-409">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-409">
         -ImageCoercion</span></span><br><span data-ttu-id="5e377-410">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-410">
         -MatrixBindings</span></span><br><span data-ttu-id="5e377-411">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-411">
         -MatrixCoercion</span></span><br><span data-ttu-id="5e377-412">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-412">
         -OoxmlCoercion</span></span><br><span data-ttu-id="5e377-413">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5e377-413">
         -PdfFile</span></span><br><span data-ttu-id="5e377-414">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="5e377-414">
         - Selection</span></span><br><span data-ttu-id="5e377-415">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-415">
         - Settings</span></span><br><span data-ttu-id="5e377-416">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-416">
         -TableBindings</span></span><br><span data-ttu-id="5e377-417">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-417">
         -TableCoercion</span></span><br><span data-ttu-id="5e377-418">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5e377-418">
         -TextBindings</span></span><br><span data-ttu-id="5e377-419">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-419">
         -TextCoercion</span></span><br><span data-ttu-id="5e377-420">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5e377-420">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="5e377-421">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5e377-421">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5e377-422">平台</span><span class="sxs-lookup"><span data-stu-id="5e377-422">Platform</span></span></th>
    <th><span data-ttu-id="5e377-423">扩展点</span><span class="sxs-lookup"><span data-stu-id="5e377-423">Extension points</span></span></th>
    <th><span data-ttu-id="5e377-424">API 要求集</span><span class="sxs-lookup"><span data-stu-id="5e377-424">API requirement sets</span></span></th>
    <th><span data-ttu-id="5e377-425"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5e377-425"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-426">Office Online</span><span class="sxs-lookup"><span data-stu-id="5e377-426">Office Online</span></span></td>
    <td> <span data-ttu-id="5e377-427">- 内容</span><span class="sxs-lookup"><span data-stu-id="5e377-427">- Content</span></span><br><span data-ttu-id="5e377-428">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-428">
         - Taskpane</span></span><br><span data-ttu-id="5e377-429">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="5e377-429">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5e377-430">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-430">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5e377-431">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5e377-431">-ActiveView</span></span><br><span data-ttu-id="5e377-432">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5e377-432">
         -CompressedFile</span></span><br><span data-ttu-id="5e377-433">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-433">
         -DocumentEvents</span></span><br><span data-ttu-id="5e377-434">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="5e377-434">
         - File</span></span><br><span data-ttu-id="5e377-435">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-435">
         -ImageCoercion</span></span><br><span data-ttu-id="5e377-436">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5e377-436">
         -PdfFile</span></span><br><span data-ttu-id="5e377-437">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="5e377-437">
         - Selection</span></span><br><span data-ttu-id="5e377-438">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-438">
         - Settings</span></span><br><span data-ttu-id="5e377-439">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-439">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-440">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="5e377-440">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="5e377-441">- 内容</span><span class="sxs-lookup"><span data-stu-id="5e377-441">- Content</span></span><br><span data-ttu-id="5e377-442">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-442">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="5e377-443">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5e377-443">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5e377-444">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5e377-444">-ActiveView</span></span><br><span data-ttu-id="5e377-445">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5e377-445">
         -CompressedFile</span></span><br><span data-ttu-id="5e377-446">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-446">
         -DocumentEvents</span></span><br><span data-ttu-id="5e377-447">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="5e377-447">
         - File</span></span><br><span data-ttu-id="5e377-448">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-448">
         -ImageCoercion</span></span><br><span data-ttu-id="5e377-449">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5e377-449">
         -PdfFile</span></span><br><span data-ttu-id="5e377-450">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="5e377-450">
         - Selection</span></span><br><span data-ttu-id="5e377-451">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-451">
         - Settings</span></span><br><span data-ttu-id="5e377-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-452">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-453">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="5e377-453">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="5e377-454">- 内容</span><span class="sxs-lookup"><span data-stu-id="5e377-454">- Content</span></span><br><span data-ttu-id="5e377-455">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-455">
         - Taskpane</span></span><br><span data-ttu-id="5e377-456">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="5e377-456">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5e377-457">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-457">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5e377-458">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5e377-458">-ActiveView</span></span><br><span data-ttu-id="5e377-459">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5e377-459">
         -CompressedFile</span></span><br><span data-ttu-id="5e377-460">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-460">
         -DocumentEvents</span></span><br><span data-ttu-id="5e377-461">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="5e377-461">
         - File</span></span><br><span data-ttu-id="5e377-462">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-462">
         -ImageCoercion</span></span><br><span data-ttu-id="5e377-463">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5e377-463">
         -PdfFile</span></span><br><span data-ttu-id="5e377-464">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="5e377-464">
         - Selection</span></span><br><span data-ttu-id="5e377-465">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-465">
         - Settings</span></span><br><span data-ttu-id="5e377-466">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-466">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-467">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="5e377-467">Office for iOS</span></span></td>
    <td> <span data-ttu-id="5e377-468">- 内容</span><span class="sxs-lookup"><span data-stu-id="5e377-468">- Content</span></span><br><span data-ttu-id="5e377-469">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-469">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="5e377-470">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-470">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="5e377-471">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5e377-471">-ActiveView</span></span><br><span data-ttu-id="5e377-472">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5e377-472">
         -CompressedFile</span></span><br><span data-ttu-id="5e377-473">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-473">
         -DocumentEvents</span></span><br><span data-ttu-id="5e377-474">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="5e377-474">
         - File</span></span><br><span data-ttu-id="5e377-475">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5e377-475">
         -PdfFile</span></span><br><span data-ttu-id="5e377-476">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="5e377-476">
         - Selection</span></span><br><span data-ttu-id="5e377-477">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-477">
         - Settings</span></span><br><span data-ttu-id="5e377-478">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-478">
         -TextCoercion</span></span><br><span data-ttu-id="5e377-479">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-479">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-480">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="5e377-480">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="5e377-481">- 内容</span><span class="sxs-lookup"><span data-stu-id="5e377-481">- Content</span></span><br><span data-ttu-id="5e377-482">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-482">
         - Taskpane</span></span><br><span data-ttu-id="5e377-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="5e377-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5e377-484">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-484">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5e377-485">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5e377-485">-ActiveView</span></span><br><span data-ttu-id="5e377-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5e377-486">
         -CompressedFile</span></span><br><span data-ttu-id="5e377-487">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-487">
         -DocumentEvents</span></span><br><span data-ttu-id="5e377-488">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="5e377-488">
         - File</span></span><br><span data-ttu-id="5e377-489">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-489">
         -ImageCoercion</span></span><br><span data-ttu-id="5e377-490">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5e377-490">
         -PdfFile</span></span><br><span data-ttu-id="5e377-491">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="5e377-491">
         - Selection</span></span><br><span data-ttu-id="5e377-492">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-492">
         - Settings</span></span><br><span data-ttu-id="5e377-493">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-493">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="5e377-494">OneNote</span><span class="sxs-lookup"><span data-stu-id="5e377-494">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5e377-495">平台</span><span class="sxs-lookup"><span data-stu-id="5e377-495">Platform</span></span></th>
    <th><span data-ttu-id="5e377-496">扩展点</span><span class="sxs-lookup"><span data-stu-id="5e377-496">Extension points</span></span></th>
    <th><span data-ttu-id="5e377-497">API 要求集</span><span class="sxs-lookup"><span data-stu-id="5e377-497">API requirement sets</span></span></th>
    <th><span data-ttu-id="5e377-498"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5e377-498"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="5e377-499">Office Online</span><span class="sxs-lookup"><span data-stu-id="5e377-499">Office Online</span></span></td>
    <td> <span data-ttu-id="5e377-500">- 内容</span><span class="sxs-lookup"><span data-stu-id="5e377-500">- Content</span></span><br><span data-ttu-id="5e377-501">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="5e377-501">
         - Taskpane</span></span><br><span data-ttu-id="5e377-502">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="5e377-502">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5e377-503">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-503">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="5e377-504">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5e377-504">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5e377-505">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5e377-505">-DocumentEvents</span></span><br><span data-ttu-id="5e377-506">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-506">
         -HtmlCoercion</span></span><br><span data-ttu-id="5e377-507">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-507">
         -ImageCoercion</span></span><br><span data-ttu-id="5e377-508">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="5e377-508">
         - Settings</span></span><br><span data-ttu-id="5e377-509">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5e377-509">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="5e377-510">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5e377-510">See also</span></span>

- [<span data-ttu-id="5e377-511">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="5e377-511">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="5e377-512">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="5e377-512">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="5e377-513">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="5e377-513">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="5e377-514">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="5e377-514">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)

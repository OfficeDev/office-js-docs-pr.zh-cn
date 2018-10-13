---
title: Office 加载项主机和平台可用性
description: Excel、Word、Outlook、PowerPoint 和 OneNote 支持的要求集。
ms.date: 10/03/2018
ms.openlocfilehash: bc7ac5c97c041a546c160c05cffc2c80db1ff1b1
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506348"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="269b8-103">Office 加载项主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="269b8-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="269b8-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表包含每个 Office 应用程序目前支持的可用平台、扩展点、API 要求集和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="269b8-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="269b8-p102">如果表格单元格内有星号 (\*)，表示我们正在完善它。有关 Project 或 Access 要求集，请参阅 [Office 通用要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="269b8-p102">If a table cell contains an asterisk ( \* ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="269b8-p103">通过 MSI 安装的 Office 2016 的内部版本号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="269b8-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="269b8-110">Excel</span><span class="sxs-lookup"><span data-stu-id="269b8-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="269b8-111">平台</span><span class="sxs-lookup"><span data-stu-id="269b8-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="269b8-112">扩展点</span><span class="sxs-lookup"><span data-stu-id="269b8-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="269b8-113">API 要求集</span><span class="sxs-lookup"><span data-stu-id="269b8-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="269b8-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="269b8-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="269b8-115">Office Online</span></span></td>
    <td> <span data-ttu-id="269b8-116">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-116">- Taskpane</span></span><br><span data-ttu-id="269b8-117">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="269b8-117">
        - Content</span></span><br><span data-ttu-id="269b8-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="269b8-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="269b8-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="269b8-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="269b8-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="269b8-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="269b8-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="269b8-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="269b8-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="269b8-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="269b8-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="269b8-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="269b8-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="269b8-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="269b8-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="269b8-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="269b8-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-128">
        -BindingEvents</span></span><br><span data-ttu-id="269b8-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-129">
        -CompressedFile</span></span><br><span data-ttu-id="269b8-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-130">
        -DocumentEvents</span></span><br><span data-ttu-id="269b8-131">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-131">
        - File</span></span><br><span data-ttu-id="269b8-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-132">
        -MatrixBindings</span></span><br><span data-ttu-id="269b8-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-133">
        -MatrixCoercion</span></span><br><span data-ttu-id="269b8-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-134">
        - Selection</span></span><br><span data-ttu-id="269b8-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-135">
        - Settings</span></span><br><span data-ttu-id="269b8-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-136">
        -TableBindings</span></span><br><span data-ttu-id="269b8-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-137">
        -TableCoercion</span></span><br><span data-ttu-id="269b8-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-138">
        -TextBindings</span></span><br><span data-ttu-id="269b8-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-139">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-140">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="269b8-140">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="269b8-141">
        - Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-141">
        - Taskpane</span></span><br><span data-ttu-id="269b8-142">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="269b8-142">
        - Content</span></span></td>
    <td>  <span data-ttu-id="269b8-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="269b8-144">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-144">
        -BindingEvents</span></span><br><span data-ttu-id="269b8-145">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-145">
        -CompressedFile</span></span><br><span data-ttu-id="269b8-146">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-146">
        -DocumentEvents</span></span><br><span data-ttu-id="269b8-147">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-147">
        - File</span></span><br><span data-ttu-id="269b8-148">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-148">
        -ImageCoercion</span></span><br><span data-ttu-id="269b8-149">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-149">
        -MatrixBindings</span></span><br><span data-ttu-id="269b8-150">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-150">
        -MatrixCoercion</span></span><br><span data-ttu-id="269b8-151">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-151">
        - Selection</span></span><br><span data-ttu-id="269b8-152">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-152">
        - Settings</span></span><br><span data-ttu-id="269b8-153">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-153">
        -TableBindings</span></span><br><span data-ttu-id="269b8-154">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-154">
        -TableCoercion</span></span><br><span data-ttu-id="269b8-155">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-155">
        -TextBindings</span></span><br><span data-ttu-id="269b8-156">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-156">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-157">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="269b8-157">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="269b8-158">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-158">- Taskpane</span></span><br><span data-ttu-id="269b8-159">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="269b8-159">
        - Content</span></span><br><span data-ttu-id="269b8-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="269b8-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="269b8-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="269b8-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="269b8-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="269b8-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="269b8-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="269b8-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="269b8-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="269b8-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="269b8-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="269b8-167">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="269b8-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="269b8-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="269b8-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="269b8-170">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-170">-BindingEvents</span></span><br><span data-ttu-id="269b8-171">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-171">
        -CompressedFile</span></span><br><span data-ttu-id="269b8-172">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-172">
        -DocumentEvents</span></span><br><span data-ttu-id="269b8-173">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-173">
        - File</span></span><br><span data-ttu-id="269b8-174">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-174">
        -ImageCoercion</span></span><br><span data-ttu-id="269b8-175">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-175">
        -MatrixBindings</span></span><br><span data-ttu-id="269b8-176">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-176">
        -MatrixCoercion</span></span><br><span data-ttu-id="269b8-177">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-177">
        - Selection</span></span><br><span data-ttu-id="269b8-178">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-178">
        - Settings</span></span><br><span data-ttu-id="269b8-179">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-179">
        -TableBindings</span></span><br><span data-ttu-id="269b8-180">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-180">
        -TableCoercion</span></span><br><span data-ttu-id="269b8-181">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-181">
        -TextBindings</span></span><br><span data-ttu-id="269b8-182">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-182">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-183">Office for Windows</span><span class="sxs-lookup"><span data-stu-id="269b8-183">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="269b8-184">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-184">- Taskpane</span></span><br><span data-ttu-id="269b8-185">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="269b8-185">
        - Content</span></span><br><span data-ttu-id="269b8-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="269b8-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="269b8-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="269b8-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="269b8-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="269b8-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="269b8-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="269b8-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="269b8-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="269b8-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="269b8-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="269b8-193">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="269b8-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="269b8-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="269b8-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="269b8-196">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-196">-BindingEvents</span></span><br><span data-ttu-id="269b8-197">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-197">
        -CompressedFile</span></span><br><span data-ttu-id="269b8-198">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-198">
        -DocumentEvents</span></span><br><span data-ttu-id="269b8-199">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-199">
        - File</span></span><br><span data-ttu-id="269b8-200">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-200">
        -ImageCoercion</span></span><br><span data-ttu-id="269b8-201">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-201">
        -MatrixBindings</span></span><br><span data-ttu-id="269b8-202">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-202">
        -MatrixCoercion</span></span><br><span data-ttu-id="269b8-203">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-203">
        - Selection</span></span><br><span data-ttu-id="269b8-204">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-204">
        - Settings</span></span><br><span data-ttu-id="269b8-205">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-205">
        -TableBindings</span></span><br><span data-ttu-id="269b8-206">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-206">
        -TableCoercion</span></span><br><span data-ttu-id="269b8-207">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-207">
        -TextBindings</span></span><br><span data-ttu-id="269b8-208">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-208">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-209">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="269b8-209">Office for iOS</span></span></td>
    <td><span data-ttu-id="269b8-210">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-210">- Taskpane</span></span><br><span data-ttu-id="269b8-211">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="269b8-211">
        - Content</span></span></td>
    <td><span data-ttu-id="269b8-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="269b8-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="269b8-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="269b8-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="269b8-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="269b8-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="269b8-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="269b8-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="269b8-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="269b8-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="269b8-218">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="269b8-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="269b8-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="269b8-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="269b8-221">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-221">-BindingEvents</span></span><br><span data-ttu-id="269b8-222">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-222">
        -CompressedFile</span></span><br><span data-ttu-id="269b8-223">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-223">
        -DocumentEvents</span></span><br><span data-ttu-id="269b8-224">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-224">
        - File</span></span><br><span data-ttu-id="269b8-225">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-225">
        -ImageCoercion</span></span><br><span data-ttu-id="269b8-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-226">
        -MatrixBindings</span></span><br><span data-ttu-id="269b8-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-227">
        -MatrixCoercion</span></span><br><span data-ttu-id="269b8-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-228">
        - Selection</span></span><br><span data-ttu-id="269b8-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-229">
        - Settings</span></span><br><span data-ttu-id="269b8-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-230">
        -TableBindings</span></span><br><span data-ttu-id="269b8-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-231">
        -TableCoercion</span></span><br><span data-ttu-id="269b8-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-232">
        -TextBindings</span></span><br><span data-ttu-id="269b8-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-233">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-234">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="269b8-234">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="269b8-235">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-235">- Taskpane</span></span><br><span data-ttu-id="269b8-236">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="269b8-236">
        - Content</span></span><br><span data-ttu-id="269b8-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="269b8-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="269b8-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="269b8-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="269b8-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="269b8-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="269b8-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="269b8-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="269b8-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="269b8-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="269b8-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="269b8-244">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="269b8-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="269b8-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="269b8-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="269b8-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-247">-BindingEvents</span></span><br><span data-ttu-id="269b8-248">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-248">
        -CompressedFile</span></span><br><span data-ttu-id="269b8-249">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-249">
        -DocumentEvents</span></span><br><span data-ttu-id="269b8-250">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-250">
        - File</span></span><br><span data-ttu-id="269b8-251">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-251">
        -ImageCoercion</span></span><br><span data-ttu-id="269b8-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-252">
        -MatrixBindings</span></span><br><span data-ttu-id="269b8-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-253">
        -MatrixCoercion</span></span><br><span data-ttu-id="269b8-254">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-254">
        -PdfFile</span></span><br><span data-ttu-id="269b8-255">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-255">
        - Selection</span></span><br><span data-ttu-id="269b8-256">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-256">
        - Settings</span></span><br><span data-ttu-id="269b8-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-257">
        -TableBindings</span></span><br><span data-ttu-id="269b8-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-258">
        -TableCoercion</span></span><br><span data-ttu-id="269b8-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-259">
        -TextBindings</span></span><br><span data-ttu-id="269b8-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-260">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-261">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="269b8-261">Office for Mac</span></span></td>
    <td><span data-ttu-id="269b8-262">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-262">- Taskpane</span></span><br><span data-ttu-id="269b8-263">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="269b8-263">
        - Content</span></span><br><span data-ttu-id="269b8-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="269b8-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="269b8-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="269b8-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="269b8-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="269b8-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="269b8-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="269b8-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="269b8-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="269b8-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="269b8-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="269b8-271">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="269b8-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="269b8-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="269b8-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="269b8-274">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-274">-BindingEvents</span></span><br><span data-ttu-id="269b8-275">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-275">
        -CompressedFile</span></span><br><span data-ttu-id="269b8-276">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-276">
        -DocumentEvents</span></span><br><span data-ttu-id="269b8-277">
        - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-277">
        - File</span></span><br><span data-ttu-id="269b8-278">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-278">
        -ImageCoercion</span></span><br><span data-ttu-id="269b8-279">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-279">
        -MatrixBindings</span></span><br><span data-ttu-id="269b8-280">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-280">
        -MatrixCoercion</span></span><br><span data-ttu-id="269b8-281">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-281">
        -PdfFile</span></span><br><span data-ttu-id="269b8-282">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-282">
        - Selection</span></span><br><span data-ttu-id="269b8-283">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-283">
        - Settings</span></span><br><span data-ttu-id="269b8-284">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-284">
        -TableBindings</span></span><br><span data-ttu-id="269b8-285">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-285">
        -TableCoercion</span></span><br><span data-ttu-id="269b8-286">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-286">
        -TextBindings</span></span><br><span data-ttu-id="269b8-287">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-287">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="269b8-288">Outlook</span><span class="sxs-lookup"><span data-stu-id="269b8-288">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="269b8-289">平台</span><span class="sxs-lookup"><span data-stu-id="269b8-289">Platform</span></span></th>
    <th><span data-ttu-id="269b8-290">扩展点</span><span class="sxs-lookup"><span data-stu-id="269b8-290">Extension points</span></span></th>
    <th><span data-ttu-id="269b8-291">API 要求集</span><span class="sxs-lookup"><span data-stu-id="269b8-291">API requirement sets</span></span></th>
    <th><span data-ttu-id="269b8-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="269b8-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-293">Office Online</span><span class="sxs-lookup"><span data-stu-id="269b8-293">Office Online</span></span></td>
    <td> <span data-ttu-id="269b8-294">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="269b8-294">- Mail Read</span></span><br><span data-ttu-id="269b8-295">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="269b8-295">
      - Mail Compose</span></span><br><span data-ttu-id="269b8-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="269b8-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="269b8-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="269b8-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="269b8-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="269b8-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="269b8-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="269b8-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="269b8-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="269b8-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="269b8-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="269b8-304">不适用</span><span class="sxs-lookup"><span data-stu-id="269b8-304">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-305">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="269b8-305">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="269b8-306">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="269b8-306">- Mail Read</span></span><br><span data-ttu-id="269b8-307">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="269b8-307">
      - Mail Compose</span></span><br><span data-ttu-id="269b8-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="269b8-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="269b8-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="269b8-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="269b8-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="269b8-313">不适用</span><span class="sxs-lookup"><span data-stu-id="269b8-313">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-314">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="269b8-314">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="269b8-315">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="269b8-315">- Mail Read</span></span><br><span data-ttu-id="269b8-316">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="269b8-316">
      - Mail Compose</span></span><br><span data-ttu-id="269b8-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="269b8-318">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="269b8-318">
      - Modules</span></span></td>
    <td> <span data-ttu-id="269b8-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="269b8-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="269b8-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="269b8-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="269b8-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="269b8-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="269b8-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="269b8-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="269b8-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="269b8-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="269b8-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="269b8-326">不适用</span><span class="sxs-lookup"><span data-stu-id="269b8-326">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-327">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="269b8-327">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="269b8-328">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="269b8-328">- Mail Read</span></span><br><span data-ttu-id="269b8-329">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="269b8-329">
      - Mail Compose</span></span><br><span data-ttu-id="269b8-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="269b8-331">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="269b8-331">
      - Modules</span></span></td>
    <td> <span data-ttu-id="269b8-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="269b8-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="269b8-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="269b8-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="269b8-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="269b8-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="269b8-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="269b8-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="269b8-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="269b8-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="269b8-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="269b8-339">不适用</span><span class="sxs-lookup"><span data-stu-id="269b8-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-340">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="269b8-340">Office for iOS</span></span></td>
    <td> <span data-ttu-id="269b8-341">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="269b8-341">- Mail Read</span></span><br><span data-ttu-id="269b8-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="269b8-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="269b8-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="269b8-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="269b8-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="269b8-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="269b8-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="269b8-348">不适用</span><span class="sxs-lookup"><span data-stu-id="269b8-348">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-349">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="269b8-349">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="269b8-350">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="269b8-350">- Mail Read</span></span><br><span data-ttu-id="269b8-351">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="269b8-351">
      - Mail Compose</span></span><br><span data-ttu-id="269b8-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="269b8-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="269b8-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="269b8-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="269b8-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="269b8-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="269b8-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="269b8-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="269b8-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="269b8-359">不适用</span><span class="sxs-lookup"><span data-stu-id="269b8-359">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-360">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="269b8-360">Office for Mac</span></span></td>
    <td> <span data-ttu-id="269b8-361">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="269b8-361">- Mail Read</span></span><br><span data-ttu-id="269b8-362">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="269b8-362">
      - Mail Compose</span></span><br><span data-ttu-id="269b8-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="269b8-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="269b8-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="269b8-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="269b8-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="269b8-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="269b8-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="269b8-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="269b8-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="269b8-370">不适用</span><span class="sxs-lookup"><span data-stu-id="269b8-370">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-371">Office for Android</span><span class="sxs-lookup"><span data-stu-id="269b8-371">Office for Android</span></span></td>
    <td> <span data-ttu-id="269b8-372">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="269b8-372">- Mail Read</span></span><br><span data-ttu-id="269b8-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-374">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-374">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="269b8-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="269b8-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="269b8-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="269b8-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="269b8-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="269b8-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="269b8-379">不适用</span><span class="sxs-lookup"><span data-stu-id="269b8-379">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="269b8-380">Word</span><span class="sxs-lookup"><span data-stu-id="269b8-380">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="269b8-381">平台</span><span class="sxs-lookup"><span data-stu-id="269b8-381">Platform</span></span></th>
    <th><span data-ttu-id="269b8-382">扩展点</span><span class="sxs-lookup"><span data-stu-id="269b8-382">Extension points</span></span></th>
    <th><span data-ttu-id="269b8-383">API 要求集</span><span class="sxs-lookup"><span data-stu-id="269b8-383">API requirement sets</span></span></th>
    <th><span data-ttu-id="269b8-384"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="269b8-384"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-385">Office Online</span><span class="sxs-lookup"><span data-stu-id="269b8-385">Office Online</span></span></td>
    <td> <span data-ttu-id="269b8-386">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-386">- Taskpane</span></span><br><span data-ttu-id="269b8-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="269b8-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="269b8-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="269b8-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="269b8-392">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-392">-BindingEvents</span></span><br><span data-ttu-id="269b8-393">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="269b8-393">
         -</span></span><br><span data-ttu-id="269b8-394">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-394">
         -DocumentEvents</span></span><br><span data-ttu-id="269b8-395">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-395">
         - File</span></span><br><span data-ttu-id="269b8-396">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-396">
         -HtmlCoercion</span></span><br><span data-ttu-id="269b8-397">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-397">
         -ImageCoercion</span></span><br><span data-ttu-id="269b8-398">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-398">
         -MatrixBindings</span></span><br><span data-ttu-id="269b8-399">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-399">
         -MatrixCoercion</span></span><br><span data-ttu-id="269b8-400">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-400">
         -OoxmlCoercion</span></span><br><span data-ttu-id="269b8-401">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-401">
         -PdfFile</span></span><br><span data-ttu-id="269b8-402">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-402">
         - Selection</span></span><br><span data-ttu-id="269b8-403">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-403">
         - Settings</span></span><br><span data-ttu-id="269b8-404">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-404">
         -TableBindings</span></span><br><span data-ttu-id="269b8-405">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-405">
         -TableCoercion</span></span><br><span data-ttu-id="269b8-406">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-406">
         -TextBindings</span></span><br><span data-ttu-id="269b8-407">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-407">
         -TextCoercion</span></span><br><span data-ttu-id="269b8-408">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="269b8-408">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-409">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="269b8-409">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="269b8-410">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-410">- Taskpane</span></span></td>
    <td> <span data-ttu-id="269b8-411">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-411">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="269b8-412">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-412">-BindingEvents</span></span><br><span data-ttu-id="269b8-413">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-413">
         -CompressedFile</span></span><br><span data-ttu-id="269b8-414">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="269b8-414">
         -</span></span><br><span data-ttu-id="269b8-415">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-415">
         -DocumentEvents</span></span><br><span data-ttu-id="269b8-416">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-416">
         - File</span></span><br><span data-ttu-id="269b8-417">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-417">
         -HtmlCoercion</span></span><br><span data-ttu-id="269b8-418">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-418">
         -ImageCoercion</span></span><br><span data-ttu-id="269b8-419">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-419">
         -MatrixBindings</span></span><br><span data-ttu-id="269b8-420">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-420">
         -MatrixCoercion</span></span><br><span data-ttu-id="269b8-421">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-421">
         -OoxmlCoercion</span></span><br><span data-ttu-id="269b8-422">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-422">
         -PdfFile</span></span><br><span data-ttu-id="269b8-423">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-423">
         - Selection</span></span><br><span data-ttu-id="269b8-424">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-424">
         - Settings</span></span><br><span data-ttu-id="269b8-425">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-425">
         -TableBindings</span></span><br><span data-ttu-id="269b8-426">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-426">
         -TableCoercion</span></span><br><span data-ttu-id="269b8-427">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-427">
         -TextBindings</span></span><br><span data-ttu-id="269b8-428">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-428">
         -TextCoercion</span></span><br><span data-ttu-id="269b8-429">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="269b8-429">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-430">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="269b8-430">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="269b8-431">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-431">- Taskpane</span></span><br><span data-ttu-id="269b8-432">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-432">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-433">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-433">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="269b8-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="269b8-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="269b8-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="269b8-437">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-437">-BindingEvents</span></span><br><span data-ttu-id="269b8-438">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-438">
         -CompressedFile</span></span><br><span data-ttu-id="269b8-439">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="269b8-439">
         -</span></span><br><span data-ttu-id="269b8-440">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-440">
         -DocumentEvents</span></span><br><span data-ttu-id="269b8-441">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-441">
         - File</span></span><br><span data-ttu-id="269b8-442">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-442">
         -HtmlCoercion</span></span><br><span data-ttu-id="269b8-443">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-443">
         -ImageCoercion</span></span><br><span data-ttu-id="269b8-444">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-444">
         -MatrixBindings</span></span><br><span data-ttu-id="269b8-445">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-445">
         -MatrixCoercion</span></span><br><span data-ttu-id="269b8-446">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-446">
         -OoxmlCoercion</span></span><br><span data-ttu-id="269b8-447">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-447">
         -PdfFile</span></span><br><span data-ttu-id="269b8-448">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-448">
         - Selection</span></span><br><span data-ttu-id="269b8-449">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-449">
         - Settings</span></span><br><span data-ttu-id="269b8-450">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-450">
         -TableBindings</span></span><br><span data-ttu-id="269b8-451">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-451">
         -TableCoercion</span></span><br><span data-ttu-id="269b8-452">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-452">
         -TextBindings</span></span><br><span data-ttu-id="269b8-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-453">
         -TextCoercion</span></span><br><span data-ttu-id="269b8-454">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="269b8-454">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-455">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="269b8-455">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="269b8-456">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-456">- Taskpane</span></span><br><span data-ttu-id="269b8-457">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-457">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-458">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-458">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="269b8-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="269b8-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="269b8-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="269b8-462">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-462">-BindingEvents</span></span><br><span data-ttu-id="269b8-463">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-463">
         -CompressedFile</span></span><br><span data-ttu-id="269b8-464">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="269b8-464">
         -</span></span><br><span data-ttu-id="269b8-465">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-465">
         -DocumentEvents</span></span><br><span data-ttu-id="269b8-466">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-466">
         - File</span></span><br><span data-ttu-id="269b8-467">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-467">
         -HtmlCoercion</span></span><br><span data-ttu-id="269b8-468">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-468">
         -ImageCoercion</span></span><br><span data-ttu-id="269b8-469">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-469">
         -MatrixBindings</span></span><br><span data-ttu-id="269b8-470">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-470">
         -MatrixCoercion</span></span><br><span data-ttu-id="269b8-471">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-471">
         -OoxmlCoercion</span></span><br><span data-ttu-id="269b8-472">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-472">
         -PdfFile</span></span><br><span data-ttu-id="269b8-473">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-473">
         - Selection</span></span><br><span data-ttu-id="269b8-474">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-474">
         - Settings</span></span><br><span data-ttu-id="269b8-475">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-475">
         -TableBindings</span></span><br><span data-ttu-id="269b8-476">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-476">
         -TableCoercion</span></span><br><span data-ttu-id="269b8-477">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-477">
         -TextBindings</span></span><br><span data-ttu-id="269b8-478">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-478">
         -TextCoercion</span></span><br><span data-ttu-id="269b8-479">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="269b8-479">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-480">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="269b8-480">Office for iOS</span></span></td>
    <td> <span data-ttu-id="269b8-481">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-481">- Taskpane</span></span></td>
    <td> <span data-ttu-id="269b8-482">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-482">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="269b8-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="269b8-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="269b8-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="269b8-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="269b8-486">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-486">-BindingEvents</span></span><br><span data-ttu-id="269b8-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-487">
         -CompressedFile</span></span><br><span data-ttu-id="269b8-488">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="269b8-488">
         -</span></span><br><span data-ttu-id="269b8-489">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-489">
         -DocumentEvents</span></span><br><span data-ttu-id="269b8-490">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-490">
         - File</span></span><br><span data-ttu-id="269b8-491">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-491">
         -HtmlCoercion</span></span><br><span data-ttu-id="269b8-492">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-492">
         -ImageCoercion</span></span><br><span data-ttu-id="269b8-493">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-493">
         -MatrixBindings</span></span><br><span data-ttu-id="269b8-494">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-494">
         -MatrixCoercion</span></span><br><span data-ttu-id="269b8-495">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-495">
         -OoxmlCoercion</span></span><br><span data-ttu-id="269b8-496">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-496">
         -PdfFile</span></span><br><span data-ttu-id="269b8-497">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-497">
         - Selection</span></span><br><span data-ttu-id="269b8-498">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-498">
         - Settings</span></span><br><span data-ttu-id="269b8-499">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-499">
         -TableBindings</span></span><br><span data-ttu-id="269b8-500">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-500">
         -TableCoercion</span></span><br><span data-ttu-id="269b8-501">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-501">
         -TextBindings</span></span><br><span data-ttu-id="269b8-502">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-502">
         -TextCoercion</span></span><br><span data-ttu-id="269b8-503">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="269b8-503">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-504">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="269b8-504">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="269b8-505">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-505">- Taskpane</span></span><br><span data-ttu-id="269b8-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="269b8-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="269b8-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="269b8-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="269b8-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="269b8-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-511">-BindingEvents</span></span><br><span data-ttu-id="269b8-512">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-512">
         -CompressedFile</span></span><br><span data-ttu-id="269b8-513">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="269b8-513">
         -</span></span><br><span data-ttu-id="269b8-514">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-514">
         -DocumentEvents</span></span><br><span data-ttu-id="269b8-515">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-515">
         - File</span></span><br><span data-ttu-id="269b8-516">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-516">
         -HtmlCoercion</span></span><br><span data-ttu-id="269b8-517">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-517">
         -ImageCoercion</span></span><br><span data-ttu-id="269b8-518">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-518">
         -MatrixBindings</span></span><br><span data-ttu-id="269b8-519">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-519">
         -MatrixCoercion</span></span><br><span data-ttu-id="269b8-520">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-520">
         -OoxmlCoercion</span></span><br><span data-ttu-id="269b8-521">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-521">
         -PdfFile</span></span><br><span data-ttu-id="269b8-522">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-522">
         - Selection</span></span><br><span data-ttu-id="269b8-523">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-523">
         - Settings</span></span><br><span data-ttu-id="269b8-524">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-524">
         -TableBindings</span></span><br><span data-ttu-id="269b8-525">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-525">
         -TableCoercion</span></span><br><span data-ttu-id="269b8-526">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-526">
         -TextBindings</span></span><br><span data-ttu-id="269b8-527">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-527">
         -TextCoercion</span></span><br><span data-ttu-id="269b8-528">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="269b8-528">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-529">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="269b8-529">Office for Mac</span></span></td>
    <td> <span data-ttu-id="269b8-530">- Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-530">- Taskpane</span></span><br><span data-ttu-id="269b8-531">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-531">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-532">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-532">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="269b8-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="269b8-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="269b8-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="269b8-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="269b8-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="269b8-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="269b8-536">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-536">-BindingEvents</span></span><br><span data-ttu-id="269b8-537">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-537">
         -CompressedFile</span></span><br><span data-ttu-id="269b8-538">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="269b8-538">
         -</span></span><br><span data-ttu-id="269b8-539">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-539">
         -DocumentEvents</span></span><br><span data-ttu-id="269b8-540">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-540">
         - File</span></span><br><span data-ttu-id="269b8-541">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-541">
         -HtmlCoercion</span></span><br><span data-ttu-id="269b8-542">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-542">
         -ImageCoercion</span></span><br><span data-ttu-id="269b8-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-543">
         -MatrixBindings</span></span><br><span data-ttu-id="269b8-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-544">
         -MatrixCoercion</span></span><br><span data-ttu-id="269b8-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-545">
         -OoxmlCoercion</span></span><br><span data-ttu-id="269b8-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-546">
         -PdfFile</span></span><br><span data-ttu-id="269b8-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-547">
         - Selection</span></span><br><span data-ttu-id="269b8-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-548">
         - Settings</span></span><br><span data-ttu-id="269b8-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-549">
         -TableBindings</span></span><br><span data-ttu-id="269b8-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-550">
         -TableCoercion</span></span><br><span data-ttu-id="269b8-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="269b8-551">
         -TextBindings</span></span><br><span data-ttu-id="269b8-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-552">
         -TextCoercion</span></span><br><span data-ttu-id="269b8-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="269b8-553">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="269b8-554">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="269b8-554">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="269b8-555">平台</span><span class="sxs-lookup"><span data-stu-id="269b8-555">Platform</span></span></th>
    <th><span data-ttu-id="269b8-556">扩展点</span><span class="sxs-lookup"><span data-stu-id="269b8-556">Extension points</span></span></th>
    <th><span data-ttu-id="269b8-557">API 要求集</span><span class="sxs-lookup"><span data-stu-id="269b8-557">API requirement sets</span></span></th>
    <th><span data-ttu-id="269b8-558"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="269b8-558"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-559">Office Online</span><span class="sxs-lookup"><span data-stu-id="269b8-559">Office Online</span></span></td>
    <td> <span data-ttu-id="269b8-560">- 内容</span><span class="sxs-lookup"><span data-stu-id="269b8-560">- Content</span></span><br><span data-ttu-id="269b8-561">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-561">
         - Taskpane</span></span><br><span data-ttu-id="269b8-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-563">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-563">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="269b8-564">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="269b8-564">-ActiveView</span></span><br><span data-ttu-id="269b8-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-565">
         -CompressedFile</span></span><br><span data-ttu-id="269b8-566">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-566">
         -DocumentEvents</span></span><br><span data-ttu-id="269b8-567">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-567">
         - File</span></span><br><span data-ttu-id="269b8-568">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-568">
         -ImageCoercion</span></span><br><span data-ttu-id="269b8-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-569">
         -PdfFile</span></span><br><span data-ttu-id="269b8-570">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-570">
         - Selection</span></span><br><span data-ttu-id="269b8-571">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-571">
         - Settings</span></span><br><span data-ttu-id="269b8-572">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-572">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-573">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="269b8-573">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="269b8-574">- 内容</span><span class="sxs-lookup"><span data-stu-id="269b8-574">- Content</span></span><br><span data-ttu-id="269b8-575">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-575">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="269b8-576">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="269b8-576">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="269b8-577">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="269b8-577">-ActiveView</span></span><br><span data-ttu-id="269b8-578">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-578">
         -CompressedFile</span></span><br><span data-ttu-id="269b8-579">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-579">
         -DocumentEvents</span></span><br><span data-ttu-id="269b8-580">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-580">
         - File</span></span><br><span data-ttu-id="269b8-581">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-581">
         -ImageCoercion</span></span><br><span data-ttu-id="269b8-582">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-582">
         -PdfFile</span></span><br><span data-ttu-id="269b8-583">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-583">
         - Selection</span></span><br><span data-ttu-id="269b8-584">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-584">
         - Settings</span></span><br><span data-ttu-id="269b8-585">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-585">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-586">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="269b8-586">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="269b8-587">- 内容</span><span class="sxs-lookup"><span data-stu-id="269b8-587">- Content</span></span><br><span data-ttu-id="269b8-588">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-588">
         - Taskpane</span></span><br><span data-ttu-id="269b8-589">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-589">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-590">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-590">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="269b8-591">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="269b8-591">-ActiveView</span></span><br><span data-ttu-id="269b8-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-592">
         -CompressedFile</span></span><br><span data-ttu-id="269b8-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-593">
         -DocumentEvents</span></span><br><span data-ttu-id="269b8-594">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-594">
         - File</span></span><br><span data-ttu-id="269b8-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-595">
         -ImageCoercion</span></span><br><span data-ttu-id="269b8-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-596">
         -PdfFile</span></span><br><span data-ttu-id="269b8-597">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-597">
         - Selection</span></span><br><span data-ttu-id="269b8-598">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-598">
         - Settings</span></span><br><span data-ttu-id="269b8-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-599">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-600">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="269b8-600">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="269b8-601">- 内容</span><span class="sxs-lookup"><span data-stu-id="269b8-601">- Content</span></span><br><span data-ttu-id="269b8-602">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-602">
         - Taskpane</span></span><br><span data-ttu-id="269b8-603">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-603">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="269b8-605">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="269b8-605">-ActiveView</span></span><br><span data-ttu-id="269b8-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-606">
         -CompressedFile</span></span><br><span data-ttu-id="269b8-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-607">
         -DocumentEvents</span></span><br><span data-ttu-id="269b8-608">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-608">
         - File</span></span><br><span data-ttu-id="269b8-609">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-609">
         -ImageCoercion</span></span><br><span data-ttu-id="269b8-610">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-610">
         -PdfFile</span></span><br><span data-ttu-id="269b8-611">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-611">
         - Selection</span></span><br><span data-ttu-id="269b8-612">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-612">
         - Settings</span></span><br><span data-ttu-id="269b8-613">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-613">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-614">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="269b8-614">Office for iOS</span></span></td>
    <td> <span data-ttu-id="269b8-615">- 内容</span><span class="sxs-lookup"><span data-stu-id="269b8-615">- Content</span></span><br><span data-ttu-id="269b8-616">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-616">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="269b8-617">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-617">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="269b8-618">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="269b8-618">-ActiveView</span></span><br><span data-ttu-id="269b8-619">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-619">
         -CompressedFile</span></span><br><span data-ttu-id="269b8-620">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-620">
         -DocumentEvents</span></span><br><span data-ttu-id="269b8-621">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-621">
         - File</span></span><br><span data-ttu-id="269b8-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-622">
         -PdfFile</span></span><br><span data-ttu-id="269b8-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-623">
         - Selection</span></span><br><span data-ttu-id="269b8-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-624">
         - Settings</span></span><br><span data-ttu-id="269b8-625">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-625">
         -TextCoercion</span></span><br><span data-ttu-id="269b8-626">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-626">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-627">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="269b8-627">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="269b8-628">- 内容</span><span class="sxs-lookup"><span data-stu-id="269b8-628">- Content</span></span><br><span data-ttu-id="269b8-629">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-629">
         - Taskpane</span></span><br><span data-ttu-id="269b8-630">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-630">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-631">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-631">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="269b8-632">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="269b8-632">-ActiveView</span></span><br><span data-ttu-id="269b8-633">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-633">
         -CompressedFile</span></span><br><span data-ttu-id="269b8-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-634">
         -DocumentEvents</span></span><br><span data-ttu-id="269b8-635">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-635">
         - File</span></span><br><span data-ttu-id="269b8-636">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-636">
         -ImageCoercion</span></span><br><span data-ttu-id="269b8-637">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-637">
         -PdfFile</span></span><br><span data-ttu-id="269b8-638">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-638">
         - Selection</span></span><br><span data-ttu-id="269b8-639">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-639">
         - Settings</span></span><br><span data-ttu-id="269b8-640">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-640">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-641">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="269b8-641">Office for Mac</span></span></td>
    <td> <span data-ttu-id="269b8-642">- 内容</span><span class="sxs-lookup"><span data-stu-id="269b8-642">- Content</span></span><br><span data-ttu-id="269b8-643">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-643">
         - Taskpane</span></span><br><span data-ttu-id="269b8-644">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-644">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-645">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-645">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="269b8-646">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="269b8-646">-ActiveView</span></span><br><span data-ttu-id="269b8-647">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="269b8-647">
         -CompressedFile</span></span><br><span data-ttu-id="269b8-648">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-648">
         -DocumentEvents</span></span><br><span data-ttu-id="269b8-649">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="269b8-649">
         - File</span></span><br><span data-ttu-id="269b8-650">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-650">
         -ImageCoercion</span></span><br><span data-ttu-id="269b8-651">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="269b8-651">
         -PdfFile</span></span><br><span data-ttu-id="269b8-652">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="269b8-652">
         - Selection</span></span><br><span data-ttu-id="269b8-653">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-653">
         - Settings</span></span><br><span data-ttu-id="269b8-654">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-654">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="269b8-655">OneNote</span><span class="sxs-lookup"><span data-stu-id="269b8-655">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="269b8-656">平台</span><span class="sxs-lookup"><span data-stu-id="269b8-656">Platform</span></span></th>
    <th><span data-ttu-id="269b8-657">扩展点</span><span class="sxs-lookup"><span data-stu-id="269b8-657">Extension points</span></span></th>
    <th><span data-ttu-id="269b8-658">API 要求集</span><span class="sxs-lookup"><span data-stu-id="269b8-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="269b8-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="269b8-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="269b8-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="269b8-660">Office Online</span></span></td>
    <td> <span data-ttu-id="269b8-661">- 内容</span><span class="sxs-lookup"><span data-stu-id="269b8-661">- Content</span></span><br><span data-ttu-id="269b8-662">
         - Taskpane</span><span class="sxs-lookup"><span data-stu-id="269b8-662">
         - Taskpane</span></span><br><span data-ttu-id="269b8-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="269b8-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="269b8-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="269b8-665">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="269b8-665">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="269b8-666">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="269b8-666">-DocumentEvents</span></span><br><span data-ttu-id="269b8-667">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-667">
         -HtmlCoercion</span></span><br><span data-ttu-id="269b8-668">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-668">
         -ImageCoercion</span></span><br><span data-ttu-id="269b8-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="269b8-669">
         - Settings</span></span><br><span data-ttu-id="269b8-670">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="269b8-670">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="269b8-671">另请参阅</span><span class="sxs-lookup"><span data-stu-id="269b8-671">See also</span></span>

- [<span data-ttu-id="269b8-672">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="269b8-672">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="269b8-673">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="269b8-673">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="269b8-674">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="269b8-674">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="269b8-675">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="269b8-675">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)

---
title: Office 加载项主机和平台可用性
description: Excel、Word、Outlook、PowerPoint 和 OneNote 支持的要求集。
ms.date: 07/31/2018
ms.openlocfilehash: 084029c0a5b70b73eaa0b3fcc180f4a813fb8b72
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703908"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="f395d-103">Office 加载项主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="f395d-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="f395d-104">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。</span><span class="sxs-lookup"><span data-stu-id="f395d-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="f395d-105">下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="f395d-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span> 

<span data-ttu-id="f395d-106">如果表格单元格内有星号 (\*)，表示我们正在完善它。</span><span class="sxs-lookup"><span data-stu-id="f395d-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="f395d-107">有关 Project 或 Access 要求集，请参阅 [Office 通用要求集](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="f395d-107">For requirement sets for Project or Access, see [Office common requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="f395d-p103">通过 MSI 安装的 Office 2016 的生成号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="f395d-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="f395d-110">Excel</span><span class="sxs-lookup"><span data-stu-id="f395d-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="f395d-111">平台</span><span class="sxs-lookup"><span data-stu-id="f395d-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="f395d-112">扩展点</span><span class="sxs-lookup"><span data-stu-id="f395d-112">Extension points</span></span></th> 
    <th style="width:20%"><span data-ttu-id="f395d-113">API 要求集</span><span class="sxs-lookup"><span data-stu-id="f395d-113">API requirement sets</span></span></th> 
    <th style="width:40%"><span data-ttu-id="f395d-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="f395d-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="f395d-115">Office Online</span></span></td>
    <td> <span data-ttu-id="f395d-116">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-116">- Taskpane</span></span><br><span data-ttu-id="f395d-117">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="f395d-117">
        - Content</span></span><br><span data-ttu-id="f395d-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="f395d-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="f395d-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f395d-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f395d-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f395d-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f395d-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f395d-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f395d-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f395d-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f395d-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f395d-124">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f395d-124">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f395d-125">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f395d-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="f395d-126">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-126">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="f395d-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-127">
        -BindingEvents</span></span><br><span data-ttu-id="f395d-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-128">
        -DocumentEvents</span></span><br><span data-ttu-id="f395d-129">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-129">
        -MatrixBindings</span></span><br><span data-ttu-id="f395d-130">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-130">
        -MatrixCoercion</span></span><br><span data-ttu-id="f395d-131">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-131">
        -TableBindings</span></span><br><span data-ttu-id="f395d-132">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-132">
        -TableCoercion</span></span><br><span data-ttu-id="f395d-133">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-133">
        -TextBindings</span></span><br><span data-ttu-id="f395d-134">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f395d-134">
        -CompressedFile</span></span><br><span data-ttu-id="f395d-135">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="f395d-135">
        - Settings</span></span><br><span data-ttu-id="f395d-136">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-136">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-137">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="f395d-137">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="f395d-138">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-138">
        - Taskpane</span></span><br><span data-ttu-id="f395d-139">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="f395d-139">
        - Content</span></span></td>
    <td>  <span data-ttu-id="f395d-140">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-140">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="f395d-141">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-141">
        -BindingEvents</span></span><br><span data-ttu-id="f395d-142">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-142">
        -DocumentEvents</span></span><br><span data-ttu-id="f395d-143">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-143">
        -MatrixBindings</span></span><br><span data-ttu-id="f395d-144">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-144">
        -MatrixCoercion</span></span><br><span data-ttu-id="f395d-145">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-145">
        -TableBindings</span></span><br><span data-ttu-id="f395d-146">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-146">
        -TableCoercion</span></span><br><span data-ttu-id="f395d-147">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-147">
        -TextBindings</span></span><br><span data-ttu-id="f395d-148">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="f395d-148">
        - Settings</span></span><br><span data-ttu-id="f395d-149">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-149">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-150">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="f395d-150">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="f395d-151">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-151">- Taskpane</span></span><br><span data-ttu-id="f395d-152">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="f395d-152">
        - Content</span></span><br><span data-ttu-id="f395d-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="f395d-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="f395d-154">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-154">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f395d-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f395d-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f395d-156">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f395d-156">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f395d-157">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f395d-157">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f395d-158">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f395d-158">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f395d-159">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f395d-159">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f395d-160">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f395d-160">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="f395d-161">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-161">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="f395d-162">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-162">-BindingEvents</span></span><br><span data-ttu-id="f395d-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-163">
        -DocumentEvents</span></span><br><span data-ttu-id="f395d-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-164">
        -MatrixBindings</span></span><br><span data-ttu-id="f395d-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-165">
        -MatrixCoercion</span></span><br><span data-ttu-id="f395d-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-166">
        -TableBindings</span></span><br><span data-ttu-id="f395d-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-167">
        -TableCoercion</span></span><br><span data-ttu-id="f395d-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-168">
        -TextBindings</span></span><br><span data-ttu-id="f395d-169">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="f395d-169">
        - Settings</span></span><br><span data-ttu-id="f395d-170">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-170">
        -TextCoercion</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-171">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="f395d-171">Office for iOS</span></span></td>
    <td><span data-ttu-id="f395d-172">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-172">- Taskpane</span></span><br><span data-ttu-id="f395d-173">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="f395d-173">
        - Content</span></span></td>
    <td><span data-ttu-id="f395d-174">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-174">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f395d-175">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f395d-175">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f395d-176">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f395d-176">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f395d-177">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f395d-177">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f395d-178">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f395d-178">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f395d-179">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f395d-179">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f395d-180">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f395d-180">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="f395d-181">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-181">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="f395d-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-182">-BindingEvents</span></span><br><span data-ttu-id="f395d-183">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-183">
        -DocumentEvents</span></span><br><span data-ttu-id="f395d-184">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-184">
        -MatrixBindings</span></span><br><span data-ttu-id="f395d-185">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-185">
        -MatrixCoercion</span></span><br><span data-ttu-id="f395d-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-186">
        -TableBindings</span></span><br><span data-ttu-id="f395d-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-187">
        -TableCoercion</span></span><br><span data-ttu-id="f395d-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-188">
        -TextBindings</span></span><br><span data-ttu-id="f395d-189">
        - 设置</span><span class="sxs-lookup"><span data-stu-id="f395d-189">
        - Settings</span></span><br><span data-ttu-id="f395d-190">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-190">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-191">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="f395d-191">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="f395d-192">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-192">- Taskpane</span></span><br><span data-ttu-id="f395d-193">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="f395d-193">
        - Content</span></span><br><span data-ttu-id="f395d-194">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="f395d-194">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="f395d-195">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-195">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f395d-196">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f395d-196">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f395d-197">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f395d-197">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f395d-198">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f395d-198">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f395d-199">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f395d-199">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f395d-200">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f395d-200">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f395d-201">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f395d-201">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="f395d-202">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-202">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="f395d-203">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-203">-BindingEvents</span></span><br><span data-ttu-id="f395d-204">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-204">
        -DocumentEvents</span></span><br><span data-ttu-id="f395d-205">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-205">
        -MatrixBindings</span></span><br><span data-ttu-id="f395d-206">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-206">
        -MatrixCoercion</span></span><br><span data-ttu-id="f395d-207">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-207">
        -TableBindings</span></span><br><span data-ttu-id="f395d-208">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-208">
        -TableCoercion</span></span><br><span data-ttu-id="f395d-209">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-209">
        -TextBindings</span></span><br><span data-ttu-id="f395d-210">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-210">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="f395d-211">Outlook</span><span class="sxs-lookup"><span data-stu-id="f395d-211">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f395d-212">平台</span><span class="sxs-lookup"><span data-stu-id="f395d-212">Platform</span></span></th>
    <th><span data-ttu-id="f395d-213">扩展点</span><span class="sxs-lookup"><span data-stu-id="f395d-213">Extension points</span></span></th> 
    <th><span data-ttu-id="f395d-214">API 要求集</span><span class="sxs-lookup"><span data-stu-id="f395d-214">API requirement sets</span></span></th> 
    <th><span data-ttu-id="f395d-215"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="f395d-215"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-216">Office Online</span><span class="sxs-lookup"><span data-stu-id="f395d-216">Office Online</span></span></td>
    <td> <span data-ttu-id="f395d-217">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="f395d-217">- Mail Read</span></span><br><span data-ttu-id="f395d-218">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="f395d-218">
      - Mail Compose</span></span><br><span data-ttu-id="f395d-219">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="f395d-219">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f395d-220">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-220">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f395d-221">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f395d-221">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f395d-222">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f395d-222">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f395d-223">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f395d-223">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f395d-224">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f395d-224">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f395d-225">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f395d-225">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="f395d-226">不适用</span><span class="sxs-lookup"><span data-stu-id="f395d-226">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-227">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="f395d-227">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="f395d-228">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="f395d-228">- Mail Read</span></span><br><span data-ttu-id="f395d-229">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="f395d-229">
      - Mail Compose</span></span><br><span data-ttu-id="f395d-230">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="f395d-230">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f395d-231">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-231">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f395d-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f395d-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f395d-233">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f395d-233">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f395d-234">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f395d-234">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="f395d-235">不适用</span><span class="sxs-lookup"><span data-stu-id="f395d-235">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-236">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="f395d-236">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="f395d-237">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="f395d-237">- Mail Read</span></span><br><span data-ttu-id="f395d-238">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="f395d-238">
      - Mail Compose</span></span><br><span data-ttu-id="f395d-239">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="f395d-239">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="f395d-240">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="f395d-240">
      - Modules</span></span></td>
    <td> <span data-ttu-id="f395d-241">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-241">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f395d-242">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f395d-242">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f395d-243">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f395d-243">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f395d-244">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f395d-244">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f395d-245">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f395d-245">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f395d-246">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f395d-246">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="f395d-247">不适用</span><span class="sxs-lookup"><span data-stu-id="f395d-247">Not available</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-248">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="f395d-248">Office for iOS</span></span></td>
    <td> <span data-ttu-id="f395d-249">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="f395d-249">- Mail Read</span></span><br><span data-ttu-id="f395d-250">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="f395d-250">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f395d-251">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-251">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f395d-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f395d-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f395d-253">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f395d-253">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f395d-254">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f395d-254">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f395d-255">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f395d-255">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>    
    <td><span data-ttu-id="f395d-256">不适用</span><span class="sxs-lookup"><span data-stu-id="f395d-256">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-257">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="f395d-257">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="f395d-258">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="f395d-258">- Mail Read</span></span><br><span data-ttu-id="f395d-259">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="f395d-259">
      - Mail Compose</span></span><br><span data-ttu-id="f395d-260">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="f395d-260">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f395d-261">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-261">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f395d-262">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f395d-262">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f395d-263">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f395d-263">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f395d-264">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f395d-264">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f395d-265">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f395d-265">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f395d-266">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f395d-266">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="f395d-267">不适用</span><span class="sxs-lookup"><span data-stu-id="f395d-267">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-268">Office for Android</span><span class="sxs-lookup"><span data-stu-id="f395d-268">Office for Android</span></span></td>
    <td> <span data-ttu-id="f395d-269">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="f395d-269">- Mail Read</span></span><br><span data-ttu-id="f395d-270">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="f395d-270">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f395d-271">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-271">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f395d-272">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f395d-272">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f395d-273">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f395d-273">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f395d-274">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f395d-274">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f395d-275">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f395d-275">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="f395d-276">不适用</span><span class="sxs-lookup"><span data-stu-id="f395d-276">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="f395d-277">Word</span><span class="sxs-lookup"><span data-stu-id="f395d-277">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f395d-278">平台</span><span class="sxs-lookup"><span data-stu-id="f395d-278">Platform</span></span></th>
    <th><span data-ttu-id="f395d-279">扩展点</span><span class="sxs-lookup"><span data-stu-id="f395d-279">Extension points</span></span></th> 
    <th><span data-ttu-id="f395d-280">API 要求集</span><span class="sxs-lookup"><span data-stu-id="f395d-280">API requirement sets</span></span></th> 
    <th><span data-ttu-id="f395d-281"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="f395d-281"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-282">Office Online</span><span class="sxs-lookup"><span data-stu-id="f395d-282">Office Online</span></span></td>
    <td> <span data-ttu-id="f395d-283">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-283">- Taskpane</span></span><br><span data-ttu-id="f395d-284">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="f395d-284">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f395d-285">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-285">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f395d-286">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f395d-286">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f395d-287">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f395d-287">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f395d-288">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-288">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f395d-289">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-289">-BindingEvents</span></span><br><span data-ttu-id="f395d-290">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f395d-290">
         -</span></span><br><span data-ttu-id="f395d-291">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-291">
         -MatrixBindings</span></span><br><span data-ttu-id="f395d-292">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-292">
         -MatrixCoercion</span></span><br><span data-ttu-id="f395d-293">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-293">
         -TableBindings</span></span><br><span data-ttu-id="f395d-294">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-294">
         -TableCoercion</span></span><br><span data-ttu-id="f395d-295">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-295">
         -TextBindings</span></span><br><span data-ttu-id="f395d-296">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-296">
         -DocumentEvents</span></span><br><span data-ttu-id="f395d-297">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f395d-297">
         -TextFile</span></span><br><span data-ttu-id="f395d-298">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-298">
         -ImageCoercion</span></span><br><span data-ttu-id="f395d-299">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="f395d-299">
         - Settings</span></span><br><span data-ttu-id="f395d-300">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-300">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-301">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="f395d-301">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="f395d-302">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-302">- Taskpane</span></span></td>
    <td> <span data-ttu-id="f395d-303">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-303">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f395d-304">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-304">-BindingEvents</span></span><br><span data-ttu-id="f395d-305">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f395d-305">
         -CompressedFile</span></span><br><span data-ttu-id="f395d-306">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="f395d-306">
         -</span></span><br><span data-ttu-id="f395d-307">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-307">
         -DocumentEvents</span></span><br><span data-ttu-id="f395d-308">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="f395d-308">
         - File</span></span><br><span data-ttu-id="f395d-309">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-309">
         -HtmlCoercion</span></span><br><span data-ttu-id="f395d-310">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-310">
         -ImageCoercion</span></span><br><span data-ttu-id="f395d-311">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-311">
         -OoxmlCoercion</span></span><br><span data-ttu-id="f395d-312">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-312">
         -TableBindings</span></span><br><span data-ttu-id="f395d-313">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-313">
         -TableCoercion</span></span><br><span data-ttu-id="f395d-314">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-314">
         -TextBindings</span></span><br><span data-ttu-id="f395d-315">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f395d-315">
         -TextFile</span></span><br><span data-ttu-id="f395d-316">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="f395d-316">
         - Settings</span></span><br><span data-ttu-id="f395d-317">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-317">
         -TextCoercion</span></span><br><span data-ttu-id="f395d-318">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-318">
         -MatrixCoercion</span></span><br><span data-ttu-id="f395d-319">
         - Matrix Bindings</span><span class="sxs-lookup"><span data-stu-id="f395d-319">
         - Matrix Bindings</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-320">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="f395d-320">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="f395d-321">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-321">- Taskpane</span></span><br><span data-ttu-id="f395d-322">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="f395d-322">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f395d-323">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-323">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f395d-324">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f395d-324">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f395d-325">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f395d-325">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f395d-326">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-326">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f395d-327">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-327">-BindingEvents</span></span><br><span data-ttu-id="f395d-328">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f395d-328">
         -CompressedFile</span></span><br><span data-ttu-id="f395d-329">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="f395d-329">
         -</span></span><br><span data-ttu-id="f395d-330">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-330">
         -DocumentEvents</span></span><br><span data-ttu-id="f395d-331">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="f395d-331">
         - File</span></span><br><span data-ttu-id="f395d-332">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-332">
         -HtmlCoercion</span></span><br><span data-ttu-id="f395d-333">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-333">
         -ImageCoercion</span></span><br><span data-ttu-id="f395d-334">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-334">
         -OoxmlCoercion</span></span><br><span data-ttu-id="f395d-335">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-335">
         -TableBindings</span></span><br><span data-ttu-id="f395d-336">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-336">
         -TableCoercion</span></span><br><span data-ttu-id="f395d-337">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-337">
         -TextBindings</span></span><br><span data-ttu-id="f395d-338">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f395d-338">
         -TextFile</span></span><br><span data-ttu-id="f395d-339">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="f395d-339">
         - Settings</span></span><br><span data-ttu-id="f395d-340">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-340">
         -TextCoercion</span></span><br><span data-ttu-id="f395d-341">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-341">
         -MatrixCoercion</span></span><br><span data-ttu-id="f395d-342">
         - Matrix Bindings</span><span class="sxs-lookup"><span data-stu-id="f395d-342">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-343">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="f395d-343">Office for iOS</span></span></td>
    <td> <span data-ttu-id="f395d-344">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-344">- Taskpane</span></span></td>
    <td> <span data-ttu-id="f395d-345">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-345">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f395d-346">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f395d-346">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f395d-347">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f395d-347">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f395d-348">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="f395d-348">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="f395d-349">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-349">-BindingEvents</span></span><br><span data-ttu-id="f395d-350">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f395d-350">
         -CompressedFile</span></span><br><span data-ttu-id="f395d-351">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="f395d-351">
         -</span></span><br><span data-ttu-id="f395d-352">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-352">
         -DocumentEvents</span></span><br><span data-ttu-id="f395d-353">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="f395d-353">
         - File</span></span><br><span data-ttu-id="f395d-354">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-354">
         -HtmlCoercion</span></span><br><span data-ttu-id="f395d-355">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-355">
         -ImageCoercion</span></span><br><span data-ttu-id="f395d-356">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-356">
         -OoxmlCoercion</span></span><br><span data-ttu-id="f395d-357">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-357">
         -TableBindings</span></span><br><span data-ttu-id="f395d-358">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-358">
         -TableCoercion</span></span><br><span data-ttu-id="f395d-359">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-359">
         -TextBindings</span></span><br><span data-ttu-id="f395d-360">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f395d-360">
         -TextFile</span></span><br><span data-ttu-id="f395d-361">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="f395d-361">
         - Settings</span></span><br><span data-ttu-id="f395d-362">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-362">
         -TextCoercion</span></span><br><span data-ttu-id="f395d-363">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-363">
         -MatrixCoercion</span></span><br><span data-ttu-id="f395d-364">
         - Matrix Bindings</span><span class="sxs-lookup"><span data-stu-id="f395d-364">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-365">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="f395d-365">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="f395d-366">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-366">- Taskpane</span></span><br><span data-ttu-id="f395d-367">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="f395d-367">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f395d-368">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-368">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f395d-369">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f395d-369">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f395d-370">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f395d-370">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f395d-371">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="f395d-371">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="f395d-372">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-372">-BindingEvents</span></span><br><span data-ttu-id="f395d-373">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f395d-373">
         -CompressedFile</span></span><br><span data-ttu-id="f395d-374">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="f395d-374">
         -</span></span><br><span data-ttu-id="f395d-375">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-375">
         -DocumentEvents</span></span><br><span data-ttu-id="f395d-376">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="f395d-376">
         - File</span></span><br><span data-ttu-id="f395d-377">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-377">
         -HtmlCoercion</span></span><br><span data-ttu-id="f395d-378">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-378">
         -ImageCoercion</span></span><br><span data-ttu-id="f395d-379">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-379">
         -OoxmlCoercion</span></span><br><span data-ttu-id="f395d-380">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-380">
         -TableBindings</span></span><br><span data-ttu-id="f395d-381">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-381">
         -TableCoercion</span></span><br><span data-ttu-id="f395d-382">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f395d-382">
         -TextBindings</span></span><br><span data-ttu-id="f395d-383">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f395d-383">
         -TextFile</span></span><br><span data-ttu-id="f395d-384">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="f395d-384">
         - Settings</span></span><br><span data-ttu-id="f395d-385">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-385">
         -TextCoercion</span></span><br><span data-ttu-id="f395d-386">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-386">
         -MatrixCoercion</span></span><br><span data-ttu-id="f395d-387">
         - Matrix Bindings</span><span class="sxs-lookup"><span data-stu-id="f395d-387">
         - Matrix Bindings</span></span> </td> 
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="f395d-388">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="f395d-388">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f395d-389">平台</span><span class="sxs-lookup"><span data-stu-id="f395d-389">Platform</span></span></th>
    <th><span data-ttu-id="f395d-390">扩展点</span><span class="sxs-lookup"><span data-stu-id="f395d-390">Extension points</span></span></th> 
    <th><span data-ttu-id="f395d-391">API 要求集</span><span class="sxs-lookup"><span data-stu-id="f395d-391">API requirement sets</span></span></th> 
    <th><span data-ttu-id="f395d-392"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="f395d-392"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-393">Office Online</span><span class="sxs-lookup"><span data-stu-id="f395d-393">Office Online</span></span></td>
    <td> <span data-ttu-id="f395d-394">- 内容</span><span class="sxs-lookup"><span data-stu-id="f395d-394">- Content</span></span><br><span data-ttu-id="f395d-395">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-395">
         - Taskpane</span></span><br><span data-ttu-id="f395d-396">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="f395d-396">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f395d-397">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-397">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f395d-398">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f395d-398">-ActiveView</span></span><br><span data-ttu-id="f395d-399">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f395d-399">
         -CompressedFile</span></span><br><span data-ttu-id="f395d-400">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="f395d-400">
         - File</span></span><br><span data-ttu-id="f395d-401">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="f395d-401">
         - Selection</span></span><br><span data-ttu-id="f395d-402">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="f395d-402">
         - Settings</span></span><br><span data-ttu-id="f395d-403">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-403">
         -TextCoercion</span></span><br><span data-ttu-id="f395d-404">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-404">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-405">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="f395d-405">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="f395d-406">- 内容</span><span class="sxs-lookup"><span data-stu-id="f395d-406">- Content</span></span><br><span data-ttu-id="f395d-407">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-407">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="f395d-408">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="f395d-408">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="f395d-409">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f395d-409">-ActiveView</span></span><br><span data-ttu-id="f395d-410">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f395d-410">
         -CompressedFile</span></span><br><span data-ttu-id="f395d-411">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-411">
         -DocumentEvents</span></span><br><span data-ttu-id="f395d-412">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="f395d-412">
         - File</span></span><br><span data-ttu-id="f395d-413">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="f395d-413">
         - Selection</span></span><br><span data-ttu-id="f395d-414">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="f395d-414">
         - Settings</span></span><br><span data-ttu-id="f395d-415">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-415">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-416">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="f395d-416">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="f395d-417">- 内容</span><span class="sxs-lookup"><span data-stu-id="f395d-417">- Content</span></span><br><span data-ttu-id="f395d-418">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-418">
         - Taskpane</span></span><br><span data-ttu-id="f395d-419">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="f395d-419">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f395d-420">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-420">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f395d-421">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f395d-421">-ActiveView</span></span><br><span data-ttu-id="f395d-422">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f395d-422">
         -CompressedFile</span></span><br><span data-ttu-id="f395d-423">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-423">
         -DocumentEvents</span></span><br><span data-ttu-id="f395d-424">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="f395d-424">
         - File</span></span><br><span data-ttu-id="f395d-425">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="f395d-425">
         - Selection</span></span><br><span data-ttu-id="f395d-426">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="f395d-426">
         - Settings</span></span><br><span data-ttu-id="f395d-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-427">
         -TextCoercion</span></span><br><span data-ttu-id="f395d-428">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-428">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-429">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="f395d-429">Office for iOS</span></span></td>
    <td> <span data-ttu-id="f395d-430">- 内容</span><span class="sxs-lookup"><span data-stu-id="f395d-430">- Content</span></span><br><span data-ttu-id="f395d-431">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-431">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="f395d-432">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-432">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="f395d-433">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f395d-433">-ActiveView</span></span><br><span data-ttu-id="f395d-434">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f395d-434">
         -CompressedFile</span></span><br><span data-ttu-id="f395d-435">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-435">
         -DocumentEvents</span></span><br><span data-ttu-id="f395d-436">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="f395d-436">
         - File</span></span><br><span data-ttu-id="f395d-437">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="f395d-437">
         - Selection</span></span><br><span data-ttu-id="f395d-438">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="f395d-438">
         - Settings</span></span><br><span data-ttu-id="f395d-439">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-439">
         -TextCoercion</span></span><br><span data-ttu-id="f395d-440">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-440">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-441">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="f395d-441">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="f395d-442">- 内容</span><span class="sxs-lookup"><span data-stu-id="f395d-442">- Content</span></span><br><span data-ttu-id="f395d-443">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-443">
         - Taskpane</span></span><br><span data-ttu-id="f395d-444">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="f395d-444">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f395d-445">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-445">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f395d-446">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f395d-446">-ActiveView</span></span><br><span data-ttu-id="f395d-447">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f395d-447">
         -CompressedFile</span></span><br><span data-ttu-id="f395d-448">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-448">
         -DocumentEvents</span></span><br><span data-ttu-id="f395d-449">
         - 文件</span><span class="sxs-lookup"><span data-stu-id="f395d-449">
         - File</span></span><br><span data-ttu-id="f395d-450">
         - 选择</span><span class="sxs-lookup"><span data-stu-id="f395d-450">
         - Selection</span></span><br><span data-ttu-id="f395d-451">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="f395d-451">
         - Settings</span></span><br><span data-ttu-id="f395d-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-452">
         -TextCoercion</span></span><br><span data-ttu-id="f395d-453">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-453">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="f395d-454">OneNote</span><span class="sxs-lookup"><span data-stu-id="f395d-454">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f395d-455">平台</span><span class="sxs-lookup"><span data-stu-id="f395d-455">Platform</span></span></th>
    <th><span data-ttu-id="f395d-456">扩展点</span><span class="sxs-lookup"><span data-stu-id="f395d-456">Extension points</span></span></th> 
    <th><span data-ttu-id="f395d-457">API 要求集</span><span class="sxs-lookup"><span data-stu-id="f395d-457">API requirement sets</span></span></th> 
    <th><span data-ttu-id="f395d-458"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="f395d-458"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="f395d-459">Office Online</span><span class="sxs-lookup"><span data-stu-id="f395d-459">Office Online</span></span></td>
    <td> <span data-ttu-id="f395d-460">- 内容</span><span class="sxs-lookup"><span data-stu-id="f395d-460">- Content</span></span><br><span data-ttu-id="f395d-461">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="f395d-461">
         - Taskpane</span></span><br><span data-ttu-id="f395d-462">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="f395d-462">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f395d-463">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-463">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="f395d-464">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f395d-464">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f395d-465">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f395d-465">-DocumentEvents</span></span><br><span data-ttu-id="f395d-466">
         - 设置</span><span class="sxs-lookup"><span data-stu-id="f395d-466">
         - Settings</span></span><br><span data-ttu-id="f395d-467">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-467">
         -TextCoercion</span></span><br><span data-ttu-id="f395d-468">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-468">
         -HtmlCoercion</span></span><br><span data-ttu-id="f395d-469">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f395d-469">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="f395d-470">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f395d-470">See also</span></span>

- [<span data-ttu-id="f395d-471">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="f395d-471">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="f395d-472">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="f395d-472">Common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="f395d-473">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="f395d-473">Add-in Commands requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="f395d-474">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="f395d-474">JavaScript API for Office reference</span></span>](https://dev.office.com/reference/add-ins/javascript-api-for-office)


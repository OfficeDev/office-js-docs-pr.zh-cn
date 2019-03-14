---
title: Office 外接程序主机和平台可用性
description: Excel、Word、Outlook、PowerPoint、OneNote 和项目支持的要求集。
ms.date: 03/07/2019
localization_priority: Priority
ms.openlocfilehash: 636c6290d8c67901beb195990593727485467460
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512879"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="7e330-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="7e330-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="7e330-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="7e330-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="7e330-p102">通过 MSI 安装的 Office 2016 的生成号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="7e330-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="7e330-108">Excel</span><span class="sxs-lookup"><span data-stu-id="7e330-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="7e330-109">平台</span><span class="sxs-lookup"><span data-stu-id="7e330-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="7e330-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="7e330-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="7e330-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="7e330-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="7e330-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7e330-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="7e330-113">Office Online</span></span></td>
    <td> <span data-ttu-id="7e330-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-114">- TaskPane</span></span><br><span data-ttu-id="7e330-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="7e330-115">
        - Content</span></span><br><span data-ttu-id="7e330-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="7e330-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="7e330-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e330-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7e330-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7e330-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e330-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7e330-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e330-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7e330-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e330-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7e330-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e330-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7e330-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7e330-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7e330-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7e330-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-126">
        - BindingEvents</span></span><br><span data-ttu-id="7e330-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-127">
        - CompressedFile</span></span><br><span data-ttu-id="7e330-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-128">
        - DocumentEvents</span></span><br><span data-ttu-id="7e330-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e330-129">
        - File</span></span><br><span data-ttu-id="7e330-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-130">
        - MatrixBindings</span></span><br><span data-ttu-id="7e330-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e330-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-132">
        - Selection</span></span><br><span data-ttu-id="7e330-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-133">
        - Settings</span></span><br><span data-ttu-id="7e330-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-134">
        - TableBindings</span></span><br><span data-ttu-id="7e330-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-135">
        - TableCoercion</span></span><br><span data-ttu-id="7e330-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-136">
        - TextBindings</span></span><br><span data-ttu-id="7e330-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-138">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="7e330-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="7e330-139">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-139">
        - TaskPane</span></span><br><span data-ttu-id="7e330-140">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="7e330-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="7e330-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7e330-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="7e330-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-142">
        - BindingEvents</span></span><br><span data-ttu-id="7e330-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-143">
        - CompressedFile</span></span><br><span data-ttu-id="7e330-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-144">
        - DocumentEvents</span></span><br><span data-ttu-id="7e330-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e330-145">
        - File</span></span><br><span data-ttu-id="7e330-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-146">
        - ImageCoercion</span></span><br><span data-ttu-id="7e330-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-147">
        - MatrixBindings</span></span><br><span data-ttu-id="7e330-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e330-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-149">
        - Selection</span></span><br><span data-ttu-id="7e330-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-150">
        - Settings</span></span><br><span data-ttu-id="7e330-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-151">
        - TableBindings</span></span><br><span data-ttu-id="7e330-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-152">
        - TableCoercion</span></span><br><span data-ttu-id="7e330-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-153">
        - TextBindings</span></span><br><span data-ttu-id="7e330-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-155">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="7e330-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="7e330-156">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-156">- TaskPane</span></span><br><span data-ttu-id="7e330-157">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="7e330-157">
        - Content</span></span></td>
    <td><span data-ttu-id="7e330-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e330-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7e330-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="7e330-160">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-160">- BindingEvents</span></span><br><span data-ttu-id="7e330-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-161">
        - CompressedFile</span></span><br><span data-ttu-id="7e330-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-162">
        - DocumentEvents</span></span><br><span data-ttu-id="7e330-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e330-163">
        - File</span></span><br><span data-ttu-id="7e330-164">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-164">
        - ImageCoercion</span></span><br><span data-ttu-id="7e330-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-165">
        - MatrixBindings</span></span><br><span data-ttu-id="7e330-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e330-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-167">
        - Selection</span></span><br><span data-ttu-id="7e330-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-168">
        - Settings</span></span><br><span data-ttu-id="7e330-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-169">
        - TableBindings</span></span><br><span data-ttu-id="7e330-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-170">
        - TableCoercion</span></span><br><span data-ttu-id="7e330-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-171">
        - TextBindings</span></span><br><span data-ttu-id="7e330-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-173">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="7e330-173">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="7e330-174">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-174">- TaskPane</span></span><br><span data-ttu-id="7e330-175">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="7e330-175">
        - Content</span></span><br><span data-ttu-id="7e330-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7e330-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e330-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7e330-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7e330-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e330-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7e330-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e330-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7e330-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e330-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7e330-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e330-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7e330-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7e330-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7e330-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7e330-186">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-186">- BindingEvents</span></span><br><span data-ttu-id="7e330-187">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-187">
        - CompressedFile</span></span><br><span data-ttu-id="7e330-188">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-188">
        - DocumentEvents</span></span><br><span data-ttu-id="7e330-189">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e330-189">
        - File</span></span><br><span data-ttu-id="7e330-190">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-190">
        - ImageCoercion</span></span><br><span data-ttu-id="7e330-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-191">
        - MatrixBindings</span></span><br><span data-ttu-id="7e330-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e330-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-193">
        - Selection</span></span><br><span data-ttu-id="7e330-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-194">
        - Settings</span></span><br><span data-ttu-id="7e330-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-195">
        - TableBindings</span></span><br><span data-ttu-id="7e330-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-196">
        - TableCoercion</span></span><br><span data-ttu-id="7e330-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-197">
        - TextBindings</span></span><br><span data-ttu-id="7e330-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-199">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="7e330-199">Office for iPad</span></span></td>
    <td><span data-ttu-id="7e330-200">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-200">- TaskPane</span></span><br><span data-ttu-id="7e330-201">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="7e330-201">
        - Content</span></span></td>
    <td><span data-ttu-id="7e330-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e330-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7e330-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7e330-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e330-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7e330-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e330-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7e330-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e330-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7e330-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e330-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7e330-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7e330-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7e330-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7e330-211">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-211">- BindingEvents</span></span><br><span data-ttu-id="7e330-212">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-212">
        - CompressedFile</span></span><br><span data-ttu-id="7e330-213">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-213">
        - DocumentEvents</span></span><br><span data-ttu-id="7e330-214">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e330-214">
        - File</span></span><br><span data-ttu-id="7e330-215">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-215">
        - ImageCoercion</span></span><br><span data-ttu-id="7e330-216">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-216">
        - MatrixBindings</span></span><br><span data-ttu-id="7e330-217">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-217">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e330-218">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-218">
        - Selection</span></span><br><span data-ttu-id="7e330-219">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-219">
        - Settings</span></span><br><span data-ttu-id="7e330-220">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-220">
        - TableBindings</span></span><br><span data-ttu-id="7e330-221">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-221">
        - TableCoercion</span></span><br><span data-ttu-id="7e330-222">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-222">
        - TextBindings</span></span><br><span data-ttu-id="7e330-223">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-223">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-224">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="7e330-224">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="7e330-225">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-225">- TaskPane</span></span><br><span data-ttu-id="7e330-226">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="7e330-226">
        - Content</span></span></td>
    <td><span data-ttu-id="7e330-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e330-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7e330-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="7e330-229">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-229">- BindingEvents</span></span><br><span data-ttu-id="7e330-230">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-230">
        - CompressedFile</span></span><br><span data-ttu-id="7e330-231">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-231">
        - DocumentEvents</span></span><br><span data-ttu-id="7e330-232">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e330-232">
        - File</span></span><br><span data-ttu-id="7e330-233">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-233">
        - ImageCoercion</span></span><br><span data-ttu-id="7e330-234">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-234">
        - MatrixBindings</span></span><br><span data-ttu-id="7e330-235">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-235">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e330-236">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-236">
        - PdfFile</span></span><br><span data-ttu-id="7e330-237">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-237">
        - Selection</span></span><br><span data-ttu-id="7e330-238">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-238">
        - Settings</span></span><br><span data-ttu-id="7e330-239">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-239">
        - TableBindings</span></span><br><span data-ttu-id="7e330-240">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-240">
        - TableCoercion</span></span><br><span data-ttu-id="7e330-241">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-241">
        - TextBindings</span></span><br><span data-ttu-id="7e330-242">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-242">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-243">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="7e330-243">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="7e330-244">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-244">- TaskPane</span></span><br><span data-ttu-id="7e330-245">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="7e330-245">
        - Content</span></span><br><span data-ttu-id="7e330-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7e330-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e330-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7e330-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7e330-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e330-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7e330-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e330-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7e330-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e330-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7e330-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e330-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7e330-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7e330-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7e330-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7e330-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-256">- BindingEvents</span></span><br><span data-ttu-id="7e330-257">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-257">
        - CompressedFile</span></span><br><span data-ttu-id="7e330-258">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-258">
        - DocumentEvents</span></span><br><span data-ttu-id="7e330-259">
        - File</span><span class="sxs-lookup"><span data-stu-id="7e330-259">
        - File</span></span><br><span data-ttu-id="7e330-260">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-260">
        - ImageCoercion</span></span><br><span data-ttu-id="7e330-261">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-261">
        - MatrixBindings</span></span><br><span data-ttu-id="7e330-262">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-262">
        - MatrixCoercion</span></span><br><span data-ttu-id="7e330-263">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-263">
        - PdfFile</span></span><br><span data-ttu-id="7e330-264">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-264">
        - Selection</span></span><br><span data-ttu-id="7e330-265">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-265">
        - Settings</span></span><br><span data-ttu-id="7e330-266">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-266">
        - TableBindings</span></span><br><span data-ttu-id="7e330-267">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-267">
        - TableCoercion</span></span><br><span data-ttu-id="7e330-268">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-268">
        - TextBindings</span></span><br><span data-ttu-id="7e330-269">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-269">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="7e330-270">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="7e330-270">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="7e330-271">Outlook</span><span class="sxs-lookup"><span data-stu-id="7e330-271">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7e330-272">平台</span><span class="sxs-lookup"><span data-stu-id="7e330-272">Platform</span></span></th>
    <th><span data-ttu-id="7e330-273">扩展点</span><span class="sxs-lookup"><span data-stu-id="7e330-273">Extension points</span></span></th>
    <th><span data-ttu-id="7e330-274">API 要求集</span><span class="sxs-lookup"><span data-stu-id="7e330-274">API requirement sets</span></span></th>
    <th><span data-ttu-id="7e330-275"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7e330-275"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-276">Office Online</span><span class="sxs-lookup"><span data-stu-id="7e330-276">Office Online</span></span></td>
    <td> <span data-ttu-id="7e330-277">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="7e330-277">- Mail Read</span></span><br><span data-ttu-id="7e330-278">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="7e330-278">
      - Mail Compose</span></span><br><span data-ttu-id="7e330-279">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-279">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e330-280">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-280">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e330-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e330-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e330-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e330-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e330-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e330-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7e330-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e330-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7e330-286">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e330-286">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7e330-287">不可用</span><span class="sxs-lookup"><span data-stu-id="7e330-287">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-288">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="7e330-288">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="7e330-289">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="7e330-289">- Mail Read</span></span><br><span data-ttu-id="7e330-290">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="7e330-290">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="7e330-291">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-291">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e330-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e330-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e330-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e330-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="7e330-295">不可用</span><span class="sxs-lookup"><span data-stu-id="7e330-295">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-296">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="7e330-296">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="7e330-297">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="7e330-297">- Mail Read</span></span><br><span data-ttu-id="7e330-298">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="7e330-298">
      - Mail Compose</span></span><br><span data-ttu-id="7e330-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7e330-300">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="7e330-300">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7e330-301">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-301">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e330-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e330-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e330-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e330-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e330-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e330-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7e330-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e330-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7e330-307">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e330-307">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7e330-308">不可用</span><span class="sxs-lookup"><span data-stu-id="7e330-308">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-309">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="7e330-309">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="7e330-310">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="7e330-310">- Mail Read</span></span><br><span data-ttu-id="7e330-311">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="7e330-311">
      - Mail Compose</span></span><br><span data-ttu-id="7e330-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7e330-313">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="7e330-313">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7e330-314">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-314">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e330-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e330-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e330-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e330-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e330-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e330-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7e330-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e330-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7e330-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7e330-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7e330-321">不可用</span><span class="sxs-lookup"><span data-stu-id="7e330-321">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-322">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="7e330-322">Office for iOS</span></span></td>
    <td> <span data-ttu-id="7e330-323">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="7e330-323">- Mail Read</span></span><br><span data-ttu-id="7e330-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e330-325">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-325">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e330-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e330-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e330-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e330-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e330-329">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e330-329">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="7e330-330">不可用</span><span class="sxs-lookup"><span data-stu-id="7e330-330">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-331">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="7e330-331">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="7e330-332">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="7e330-332">- Mail Read</span></span><br><span data-ttu-id="7e330-333">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="7e330-333">
      - Mail Compose</span></span><br><span data-ttu-id="7e330-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e330-335">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-335">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e330-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e330-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e330-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e330-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e330-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e330-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7e330-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e330-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7e330-341">不可用</span><span class="sxs-lookup"><span data-stu-id="7e330-341">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-342">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="7e330-342">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="7e330-343">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="7e330-343">- Mail Read</span></span><br><span data-ttu-id="7e330-344">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="7e330-344">
      - Mail Compose</span></span><br><span data-ttu-id="7e330-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e330-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e330-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e330-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e330-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e330-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e330-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e330-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7e330-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e330-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7e330-352">不可用</span><span class="sxs-lookup"><span data-stu-id="7e330-352">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-353">Office for Android</span><span class="sxs-lookup"><span data-stu-id="7e330-353">Office for Android</span></span></td>
    <td> <span data-ttu-id="7e330-354">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="7e330-354">- Mail Read</span></span><br><span data-ttu-id="7e330-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e330-356">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-356">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e330-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e330-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e330-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e330-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e330-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e330-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="7e330-361">不可用</span><span class="sxs-lookup"><span data-stu-id="7e330-361">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="7e330-362">Word</span><span class="sxs-lookup"><span data-stu-id="7e330-362">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7e330-363">平台</span><span class="sxs-lookup"><span data-stu-id="7e330-363">Platform</span></span></th>
    <th><span data-ttu-id="7e330-364">扩展点</span><span class="sxs-lookup"><span data-stu-id="7e330-364">Extension points</span></span></th>
    <th><span data-ttu-id="7e330-365">API 要求集</span><span class="sxs-lookup"><span data-stu-id="7e330-365">API requirement sets</span></span></th>
    <th><span data-ttu-id="7e330-366"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7e330-366"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-367">Office Online</span><span class="sxs-lookup"><span data-stu-id="7e330-367">Office Online</span></span></td>
    <td> <span data-ttu-id="7e330-368">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-368">- TaskPane</span></span><br><span data-ttu-id="7e330-369">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-369">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e330-370">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-370">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7e330-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7e330-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7e330-373">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-373">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e330-374">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-374">- BindingEvents</span></span><br><span data-ttu-id="7e330-375">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e330-375">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e330-376">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-376">
         - DocumentEvents</span></span><br><span data-ttu-id="7e330-377">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e330-377">
         - File</span></span><br><span data-ttu-id="7e330-378">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-378">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e330-379">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-379">
         - ImageCoercion</span></span><br><span data-ttu-id="7e330-380">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-380">
         - MatrixBindings</span></span><br><span data-ttu-id="7e330-381">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-381">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e330-382">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-382">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e330-383">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-383">
         - PdfFile</span></span><br><span data-ttu-id="7e330-384">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-384">
         - Selection</span></span><br><span data-ttu-id="7e330-385">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-385">
         - Settings</span></span><br><span data-ttu-id="7e330-386">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-386">
         - TableBindings</span></span><br><span data-ttu-id="7e330-387">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-387">
         - TableCoercion</span></span><br><span data-ttu-id="7e330-388">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-388">
         - TextBindings</span></span><br><span data-ttu-id="7e330-389">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-389">
         - TextCoercion</span></span><br><span data-ttu-id="7e330-390">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e330-390">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-391">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="7e330-391">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="7e330-392">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-392">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7e330-393">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7e330-393">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="7e330-394">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-394">- BindingEvents</span></span><br><span data-ttu-id="7e330-395">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-395">
         - CompressedFile</span></span><br><span data-ttu-id="7e330-396">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e330-396">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e330-397">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-397">
         - DocumentEvents</span></span><br><span data-ttu-id="7e330-398">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e330-398">
         - File</span></span><br><span data-ttu-id="7e330-399">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-399">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e330-400">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-400">
         - ImageCoercion</span></span><br><span data-ttu-id="7e330-401">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-401">
         - MatrixBindings</span></span><br><span data-ttu-id="7e330-402">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-402">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e330-403">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-403">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e330-404">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-404">
         - PdfFile</span></span><br><span data-ttu-id="7e330-405">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-405">
         - Selection</span></span><br><span data-ttu-id="7e330-406">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-406">
         - Settings</span></span><br><span data-ttu-id="7e330-407">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-407">
         - TableBindings</span></span><br><span data-ttu-id="7e330-408">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-408">
         - TableCoercion</span></span><br><span data-ttu-id="7e330-409">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-409">
         - TextBindings</span></span><br><span data-ttu-id="7e330-410">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-410">
         - TextCoercion</span></span><br><span data-ttu-id="7e330-411">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e330-411">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-412">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="7e330-412">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="7e330-413">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-413">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7e330-414">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-414">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7e330-415">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7e330-415">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="7e330-416">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-416">- BindingEvents</span></span><br><span data-ttu-id="7e330-417">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-417">
         - CompressedFile</span></span><br><span data-ttu-id="7e330-418">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e330-418">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e330-419">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-419">
         - DocumentEvents</span></span><br><span data-ttu-id="7e330-420">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e330-420">
         - File</span></span><br><span data-ttu-id="7e330-421">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-421">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e330-422">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-422">
         - ImageCoercion</span></span><br><span data-ttu-id="7e330-423">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-423">
         - MatrixBindings</span></span><br><span data-ttu-id="7e330-424">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-424">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e330-425">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-425">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e330-426">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-426">
         - PdfFile</span></span><br><span data-ttu-id="7e330-427">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-427">
         - Selection</span></span><br><span data-ttu-id="7e330-428">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-428">
         - Settings</span></span><br><span data-ttu-id="7e330-429">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-429">
         - TableBindings</span></span><br><span data-ttu-id="7e330-430">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-430">
         - TableCoercion</span></span><br><span data-ttu-id="7e330-431">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-431">
         - TextBindings</span></span><br><span data-ttu-id="7e330-432">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-432">
         - TextCoercion</span></span><br><span data-ttu-id="7e330-433">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e330-433">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-434">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="7e330-434">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="7e330-435">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-435">- TaskPane</span></span><br><span data-ttu-id="7e330-436">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-436">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e330-437">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-437">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7e330-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7e330-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7e330-440">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-440">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e330-441">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-441">- BindingEvents</span></span><br><span data-ttu-id="7e330-442">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-442">
         - CompressedFile</span></span><br><span data-ttu-id="7e330-443">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e330-443">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e330-444">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-444">
         - DocumentEvents</span></span><br><span data-ttu-id="7e330-445">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e330-445">
         - File</span></span><br><span data-ttu-id="7e330-446">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-446">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e330-447">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-447">
         - ImageCoercion</span></span><br><span data-ttu-id="7e330-448">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-448">
         - MatrixBindings</span></span><br><span data-ttu-id="7e330-449">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-449">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e330-450">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-450">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e330-451">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-451">
         - PdfFile</span></span><br><span data-ttu-id="7e330-452">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-452">
         - Selection</span></span><br><span data-ttu-id="7e330-453">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-453">
         - Settings</span></span><br><span data-ttu-id="7e330-454">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-454">
         - TableBindings</span></span><br><span data-ttu-id="7e330-455">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-455">
         - TableCoercion</span></span><br><span data-ttu-id="7e330-456">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-456">
         - TextBindings</span></span><br><span data-ttu-id="7e330-457">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-457">
         - TextCoercion</span></span><br><span data-ttu-id="7e330-458">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e330-458">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-459">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="7e330-459">Office for iPad</span></span></td>
    <td> <span data-ttu-id="7e330-460">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-460">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7e330-461">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-461">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7e330-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7e330-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7e330-464">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="7e330-464">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="7e330-465">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-465">- BindingEvents</span></span><br><span data-ttu-id="7e330-466">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-466">
         - CompressedFile</span></span><br><span data-ttu-id="7e330-467">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e330-467">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e330-468">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-468">
         - DocumentEvents</span></span><br><span data-ttu-id="7e330-469">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e330-469">
         - File</span></span><br><span data-ttu-id="7e330-470">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-470">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e330-471">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-471">
         - ImageCoercion</span></span><br><span data-ttu-id="7e330-472">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-472">
         - MatrixBindings</span></span><br><span data-ttu-id="7e330-473">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-473">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e330-474">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-474">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e330-475">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-475">
         - PdfFile</span></span><br><span data-ttu-id="7e330-476">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-476">
         - Selection</span></span><br><span data-ttu-id="7e330-477">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-477">
         - Settings</span></span><br><span data-ttu-id="7e330-478">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-478">
         - TableBindings</span></span><br><span data-ttu-id="7e330-479">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-479">
         - TableCoercion</span></span><br><span data-ttu-id="7e330-480">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-480">
         - TextBindings</span></span><br><span data-ttu-id="7e330-481">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-481">
         - TextCoercion</span></span><br><span data-ttu-id="7e330-482">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e330-482">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-483">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="7e330-483">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="7e330-484">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-484">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7e330-485">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-485">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7e330-486">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7e330-486">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="7e330-487">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-487">- BindingEvents</span></span><br><span data-ttu-id="7e330-488">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-488">
         - CompressedFile</span></span><br><span data-ttu-id="7e330-489">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e330-489">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e330-490">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-490">
         - DocumentEvents</span></span><br><span data-ttu-id="7e330-491">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e330-491">
         - File</span></span><br><span data-ttu-id="7e330-492">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-492">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e330-493">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-493">
         - ImageCoercion</span></span><br><span data-ttu-id="7e330-494">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-494">
         - MatrixBindings</span></span><br><span data-ttu-id="7e330-495">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-495">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e330-496">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-496">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e330-497">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-497">
         - PdfFile</span></span><br><span data-ttu-id="7e330-498">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-498">
         - Selection</span></span><br><span data-ttu-id="7e330-499">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-499">
         - Settings</span></span><br><span data-ttu-id="7e330-500">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-500">
         - TableBindings</span></span><br><span data-ttu-id="7e330-501">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-501">
         - TableCoercion</span></span><br><span data-ttu-id="7e330-502">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-502">
         - TextBindings</span></span><br><span data-ttu-id="7e330-503">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-503">
         - TextCoercion</span></span><br><span data-ttu-id="7e330-504">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e330-504">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-505">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="7e330-505">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="7e330-506">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-506">- TaskPane</span></span><br><span data-ttu-id="7e330-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e330-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7e330-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e330-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7e330-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e330-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7e330-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="7e330-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="7e330-512">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-512">- BindingEvents</span></span><br><span data-ttu-id="7e330-513">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-513">
         - CompressedFile</span></span><br><span data-ttu-id="7e330-514">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e330-514">
         - CustomXmlParts</span></span><br><span data-ttu-id="7e330-515">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-515">
         - DocumentEvents</span></span><br><span data-ttu-id="7e330-516">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e330-516">
         - File</span></span><br><span data-ttu-id="7e330-517">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-517">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e330-518">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-518">
         - ImageCoercion</span></span><br><span data-ttu-id="7e330-519">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-519">
         - MatrixBindings</span></span><br><span data-ttu-id="7e330-520">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-520">
         - MatrixCoercion</span></span><br><span data-ttu-id="7e330-521">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-521">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7e330-522">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-522">
         - PdfFile</span></span><br><span data-ttu-id="7e330-523">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-523">
         - Selection</span></span><br><span data-ttu-id="7e330-524">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-524">
         - Settings</span></span><br><span data-ttu-id="7e330-525">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-525">
         - TableBindings</span></span><br><span data-ttu-id="7e330-526">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-526">
         - TableCoercion</span></span><br><span data-ttu-id="7e330-527">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e330-527">
         - TextBindings</span></span><br><span data-ttu-id="7e330-528">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-528">
         - TextCoercion</span></span><br><span data-ttu-id="7e330-529">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e330-529">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="7e330-530">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="7e330-530">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="7e330-531">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="7e330-531">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7e330-532">平台</span><span class="sxs-lookup"><span data-stu-id="7e330-532">Platform</span></span></th>
    <th><span data-ttu-id="7e330-533">扩展点</span><span class="sxs-lookup"><span data-stu-id="7e330-533">Extension points</span></span></th>
    <th><span data-ttu-id="7e330-534">API 要求集</span><span class="sxs-lookup"><span data-stu-id="7e330-534">API requirement sets</span></span></th>
    <th><span data-ttu-id="7e330-535"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7e330-535"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-536">Office Online</span><span class="sxs-lookup"><span data-stu-id="7e330-536">Office Online</span></span></td>
    <td> <span data-ttu-id="7e330-537">- 内容</span><span class="sxs-lookup"><span data-stu-id="7e330-537">- Content</span></span><br><span data-ttu-id="7e330-538">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-538">
         - TaskPane</span></span><br><span data-ttu-id="7e330-539">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-539">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e330-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e330-541">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e330-541">- ActiveView</span></span><br><span data-ttu-id="7e330-542">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-542">
         - CompressedFile</span></span><br><span data-ttu-id="7e330-543">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-543">
         - DocumentEvents</span></span><br><span data-ttu-id="7e330-544">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e330-544">
         - File</span></span><br><span data-ttu-id="7e330-545">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-545">
         - ImageCoercion</span></span><br><span data-ttu-id="7e330-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-546">
         - PdfFile</span></span><br><span data-ttu-id="7e330-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-547">
         - Selection</span></span><br><span data-ttu-id="7e330-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-548">
         - Settings</span></span><br><span data-ttu-id="7e330-549">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-549">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-550">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="7e330-550">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="7e330-551">- 内容</span><span class="sxs-lookup"><span data-stu-id="7e330-551">- Content</span></span><br><span data-ttu-id="7e330-552">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-552">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="7e330-553">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7e330-553">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="7e330-554">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e330-554">- ActiveView</span></span><br><span data-ttu-id="7e330-555">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-555">
         - CompressedFile</span></span><br><span data-ttu-id="7e330-556">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-556">
         - DocumentEvents</span></span><br><span data-ttu-id="7e330-557">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e330-557">
         - File</span></span><br><span data-ttu-id="7e330-558">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-558">
         - ImageCoercion</span></span><br><span data-ttu-id="7e330-559">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-559">
         - PdfFile</span></span><br><span data-ttu-id="7e330-560">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-560">
         - Selection</span></span><br><span data-ttu-id="7e330-561">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-561">
         - Settings</span></span><br><span data-ttu-id="7e330-562">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-562">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-563">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="7e330-563">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="7e330-564">- 内容</span><span class="sxs-lookup"><span data-stu-id="7e330-564">- Content</span></span><br><span data-ttu-id="7e330-565">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-565">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="7e330-566">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7e330-566">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="7e330-567">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e330-567">- ActiveView</span></span><br><span data-ttu-id="7e330-568">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-568">
         - CompressedFile</span></span><br><span data-ttu-id="7e330-569">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-569">
         - DocumentEvents</span></span><br><span data-ttu-id="7e330-570">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e330-570">
         - File</span></span><br><span data-ttu-id="7e330-571">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-571">
         - ImageCoercion</span></span><br><span data-ttu-id="7e330-572">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-572">
         - PdfFile</span></span><br><span data-ttu-id="7e330-573">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-573">
         - Selection</span></span><br><span data-ttu-id="7e330-574">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-574">
         - Settings</span></span><br><span data-ttu-id="7e330-575">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-575">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-576">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="7e330-576">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="7e330-577">- 内容</span><span class="sxs-lookup"><span data-stu-id="7e330-577">- Content</span></span><br><span data-ttu-id="7e330-578">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-578">
         - TaskPane</span></span><br><span data-ttu-id="7e330-579">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-579">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e330-580">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-580">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e330-581">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e330-581">- ActiveView</span></span><br><span data-ttu-id="7e330-582">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-582">
         - CompressedFile</span></span><br><span data-ttu-id="7e330-583">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-583">
         - DocumentEvents</span></span><br><span data-ttu-id="7e330-584">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e330-584">
         - File</span></span><br><span data-ttu-id="7e330-585">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-585">
         - ImageCoercion</span></span><br><span data-ttu-id="7e330-586">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-586">
         - PdfFile</span></span><br><span data-ttu-id="7e330-587">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-587">
         - Selection</span></span><br><span data-ttu-id="7e330-588">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-588">
         - Settings</span></span><br><span data-ttu-id="7e330-589">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-589">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-590">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="7e330-590">Office for iPad</span></span></td>
    <td> <span data-ttu-id="7e330-591">- 内容</span><span class="sxs-lookup"><span data-stu-id="7e330-591">- Content</span></span><br><span data-ttu-id="7e330-592">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-592">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="7e330-593">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-593">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="7e330-594">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e330-594">- ActiveView</span></span><br><span data-ttu-id="7e330-595">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-595">
         - CompressedFile</span></span><br><span data-ttu-id="7e330-596">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-596">
         - DocumentEvents</span></span><br><span data-ttu-id="7e330-597">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e330-597">
         - File</span></span><br><span data-ttu-id="7e330-598">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-598">
         - PdfFile</span></span><br><span data-ttu-id="7e330-599">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-599">
         - Selection</span></span><br><span data-ttu-id="7e330-600">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-600">
         - Settings</span></span><br><span data-ttu-id="7e330-601">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-601">
         - TextCoercion</span></span><br><span data-ttu-id="7e330-602">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-602">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-603">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="7e330-603">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="7e330-604">- 内容</span><span class="sxs-lookup"><span data-stu-id="7e330-604">- Content</span></span><br><span data-ttu-id="7e330-605">
         - 任务窗格/td></span><span class="sxs-lookup"><span data-stu-id="7e330-605">
         - TaskPane/td></span></span> <td> <span data-ttu-id="7e330-606">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7e330-606">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="7e330-607">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e330-607">- ActiveView</span></span><br><span data-ttu-id="7e330-608">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-608">
         - CompressedFile</span></span><br><span data-ttu-id="7e330-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-609">
         - DocumentEvents</span></span><br><span data-ttu-id="7e330-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e330-610">
         - File</span></span><br><span data-ttu-id="7e330-611">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-611">
         - ImageCoercion</span></span><br><span data-ttu-id="7e330-612">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-612">
         - PdfFile</span></span><br><span data-ttu-id="7e330-613">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-613">
         - Selection</span></span><br><span data-ttu-id="7e330-614">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-614">
         - Settings</span></span><br><span data-ttu-id="7e330-615">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-615">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-616">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="7e330-616">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="7e330-617">- 内容</span><span class="sxs-lookup"><span data-stu-id="7e330-617">- Content</span></span><br><span data-ttu-id="7e330-618">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-618">
         - TaskPane</span></span><br><span data-ttu-id="7e330-619">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-619">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e330-620">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-620">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e330-621">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e330-621">- ActiveView</span></span><br><span data-ttu-id="7e330-622">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e330-622">
         - CompressedFile</span></span><br><span data-ttu-id="7e330-623">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-623">
         - DocumentEvents</span></span><br><span data-ttu-id="7e330-624">
         - File</span><span class="sxs-lookup"><span data-stu-id="7e330-624">
         - File</span></span><br><span data-ttu-id="7e330-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-625">
         - ImageCoercion</span></span><br><span data-ttu-id="7e330-626">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7e330-626">
         - PdfFile</span></span><br><span data-ttu-id="7e330-627">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-627">
         - Selection</span></span><br><span data-ttu-id="7e330-628">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-628">
         - Settings</span></span><br><span data-ttu-id="7e330-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-629">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="7e330-630">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="7e330-630">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="7e330-631">OneNote</span><span class="sxs-lookup"><span data-stu-id="7e330-631">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7e330-632">平台</span><span class="sxs-lookup"><span data-stu-id="7e330-632">Platform</span></span></th>
    <th><span data-ttu-id="7e330-633">扩展点</span><span class="sxs-lookup"><span data-stu-id="7e330-633">Extension points</span></span></th>
    <th><span data-ttu-id="7e330-634">API 要求集</span><span class="sxs-lookup"><span data-stu-id="7e330-634">API requirement sets</span></span></th>
    <th><span data-ttu-id="7e330-635"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7e330-635"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-636">Office Online</span><span class="sxs-lookup"><span data-stu-id="7e330-636">Office Online</span></span></td>
    <td> <span data-ttu-id="7e330-637">- 内容</span><span class="sxs-lookup"><span data-stu-id="7e330-637">- Content</span></span><br><span data-ttu-id="7e330-638">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-638">
         - TaskPane</span></span><br><span data-ttu-id="7e330-639">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="7e330-639">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e330-640">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-640">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="7e330-641">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-641">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e330-642">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e330-642">- DocumentEvents</span></span><br><span data-ttu-id="7e330-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="7e330-644">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-644">
         - ImageCoercion</span></span><br><span data-ttu-id="7e330-645">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7e330-645">
         - Settings</span></span><br><span data-ttu-id="7e330-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-646">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="7e330-647">项目</span><span class="sxs-lookup"><span data-stu-id="7e330-647">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7e330-648">平台</span><span class="sxs-lookup"><span data-stu-id="7e330-648">Platform</span></span></th>
    <th><span data-ttu-id="7e330-649">扩展点</span><span class="sxs-lookup"><span data-stu-id="7e330-649">Extension points</span></span></th>
    <th><span data-ttu-id="7e330-650">API 要求集</span><span class="sxs-lookup"><span data-stu-id="7e330-650">API requirement sets</span></span></th>
    <th><span data-ttu-id="7e330-651"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7e330-651"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-652">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="7e330-652">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="7e330-653">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-653">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7e330-654">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-654">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e330-655">- Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-655">- Selection</span></span><br><span data-ttu-id="7e330-656">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-656">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-657">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="7e330-657">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="7e330-658">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-658">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7e330-659">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-659">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e330-660">- Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-660">- Selection</span></span><br><span data-ttu-id="7e330-661">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-661">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e330-662">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="7e330-662">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="7e330-663">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="7e330-663">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7e330-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e330-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e330-665">- Selection</span><span class="sxs-lookup"><span data-stu-id="7e330-665">- Selection</span></span><br><span data-ttu-id="7e330-666">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e330-666">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="7e330-667">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7e330-667">See also</span></span>

- [<span data-ttu-id="7e330-668">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="7e330-668">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="7e330-669">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="7e330-669">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="7e330-670">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="7e330-670">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="7e330-671">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="7e330-671">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)

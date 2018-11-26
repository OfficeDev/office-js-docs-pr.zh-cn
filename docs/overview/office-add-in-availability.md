---
title: Office 外接程序主机和平台可用性
description: Excel、Word、Outlook、PowerPoint、OneNote 和项目支持的要求集。
ms.date: 11/07/2018
ms.openlocfilehash: c3da40be21c0e569028dd10e93e33760ba2bd39d
ms.sourcegitcommit: 3e84d616e69f39eeeeea773f2431e7d674c4a9f5
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/22/2018
ms.locfileid: "26644751"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="e462f-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="e462f-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="e462f-104">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。</span><span class="sxs-lookup"><span data-stu-id="e462f-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="e462f-105">下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="e462f-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="e462f-p102">通过 MSI 安装的 Office 2016 的生成号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="e462f-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="e462f-108">Excel</span><span class="sxs-lookup"><span data-stu-id="e462f-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="e462f-109">平台</span><span class="sxs-lookup"><span data-stu-id="e462f-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="e462f-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="e462f-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="e462f-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="e462f-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="e462f-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e462f-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="e462f-113">Office Online</span></span></td>
    <td> <span data-ttu-id="e462f-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-114">- TaskPane</span></span><br><span data-ttu-id="e462f-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="e462f-115">
        - Content</span></span><br><span data-ttu-id="e462f-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="e462f-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="e462f-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e462f-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e462f-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e462f-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e462f-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e462f-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e462f-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e462f-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e462f-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e462f-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e462f-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e462f-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e462f-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e462f-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e462f-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-126">
        - BindingEvents</span></span><br><span data-ttu-id="e462f-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-127">
        - CompressedFile</span></span><br><span data-ttu-id="e462f-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-128">
        - DocumentEvents</span></span><br><span data-ttu-id="e462f-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="e462f-129">
        - File</span></span><br><span data-ttu-id="e462f-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-130">
        - MatrixBindings</span></span><br><span data-ttu-id="e462f-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="e462f-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-132">
        - Selection</span></span><br><span data-ttu-id="e462f-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-133">
        - Settings</span></span><br><span data-ttu-id="e462f-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-134">
        - TableBindings</span></span><br><span data-ttu-id="e462f-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-135">
        - TableCoercion</span></span><br><span data-ttu-id="e462f-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-136">
        - TextBindings</span></span><br><span data-ttu-id="e462f-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-138">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="e462f-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="e462f-139">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-139">
        - TaskPane</span></span><br><span data-ttu-id="e462f-140">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="e462f-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="e462f-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e462f-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-142">
        - BindingEvents</span></span><br><span data-ttu-id="e462f-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-143">
        - CompressedFile</span></span><br><span data-ttu-id="e462f-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-144">
        - DocumentEvents</span></span><br><span data-ttu-id="e462f-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="e462f-145">
        - File</span></span><br><span data-ttu-id="e462f-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-146">
        - ImageCoercion</span></span><br><span data-ttu-id="e462f-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-147">
        - MatrixBindings</span></span><br><span data-ttu-id="e462f-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="e462f-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-149">
        - Selection</span></span><br><span data-ttu-id="e462f-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-150">
        - Settings</span></span><br><span data-ttu-id="e462f-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-151">
        - TableBindings</span></span><br><span data-ttu-id="e462f-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-152">
        - TableCoercion</span></span><br><span data-ttu-id="e462f-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-153">
        - TextBindings</span></span><br><span data-ttu-id="e462f-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-155">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="e462f-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="e462f-156">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-156">- TaskPane</span></span><br><span data-ttu-id="e462f-157">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="e462f-157">
        - Content</span></span><br><span data-ttu-id="e462f-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e462f-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e462f-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e462f-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e462f-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e462f-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e462f-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e462f-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e462f-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e462f-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e462f-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e462f-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e462f-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e462f-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e462f-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e462f-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-168">- BindingEvents</span></span><br><span data-ttu-id="e462f-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-169">
        - CompressedFile</span></span><br><span data-ttu-id="e462f-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-170">
        - DocumentEvents</span></span><br><span data-ttu-id="e462f-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="e462f-171">
        - File</span></span><br><span data-ttu-id="e462f-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-172">
        - ImageCoercion</span></span><br><span data-ttu-id="e462f-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-173">
        - MatrixBindings</span></span><br><span data-ttu-id="e462f-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="e462f-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-175">
        - Selection</span></span><br><span data-ttu-id="e462f-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-176">
        - Settings</span></span><br><span data-ttu-id="e462f-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-177">
        - TableBindings</span></span><br><span data-ttu-id="e462f-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-178">
        - TableCoercion</span></span><br><span data-ttu-id="e462f-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-179">
        - TextBindings</span></span><br><span data-ttu-id="e462f-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-181">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="e462f-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="e462f-182">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-182">- TaskPane</span></span><br><span data-ttu-id="e462f-183">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="e462f-183">
        - Content</span></span><br><span data-ttu-id="e462f-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e462f-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e462f-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e462f-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e462f-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e462f-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e462f-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e462f-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e462f-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e462f-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e462f-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e462f-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e462f-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e462f-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e462f-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e462f-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-194">- BindingEvents</span></span><br><span data-ttu-id="e462f-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-195">
        - CompressedFile</span></span><br><span data-ttu-id="e462f-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-196">
        - DocumentEvents</span></span><br><span data-ttu-id="e462f-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="e462f-197">
        - File</span></span><br><span data-ttu-id="e462f-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-198">
        - ImageCoercion</span></span><br><span data-ttu-id="e462f-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-199">
        - MatrixBindings</span></span><br><span data-ttu-id="e462f-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="e462f-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-201">
        - Selection</span></span><br><span data-ttu-id="e462f-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-202">
        - Settings</span></span><br><span data-ttu-id="e462f-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-203">
        - TableBindings</span></span><br><span data-ttu-id="e462f-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-204">
        - TableCoercion</span></span><br><span data-ttu-id="e462f-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-205">
        - TextBindings</span></span><br><span data-ttu-id="e462f-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-207">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="e462f-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="e462f-208">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-208">- TaskPane</span></span><br><span data-ttu-id="e462f-209">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="e462f-209">
        - Content</span></span></td>
    <td><span data-ttu-id="e462f-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e462f-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e462f-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e462f-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e462f-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e462f-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e462f-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e462f-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e462f-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e462f-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e462f-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e462f-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e462f-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e462f-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e462f-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-219">- BindingEvents</span></span><br><span data-ttu-id="e462f-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-220">
        - CompressedFile</span></span><br><span data-ttu-id="e462f-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-221">
        - DocumentEvents</span></span><br><span data-ttu-id="e462f-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="e462f-222">
        - File</span></span><br><span data-ttu-id="e462f-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-223">
        - ImageCoercion</span></span><br><span data-ttu-id="e462f-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-224">
        - MatrixBindings</span></span><br><span data-ttu-id="e462f-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="e462f-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-226">
        - Selection</span></span><br><span data-ttu-id="e462f-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-227">
        - Settings</span></span><br><span data-ttu-id="e462f-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-228">
        - TableBindings</span></span><br><span data-ttu-id="e462f-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-229">
        - TableCoercion</span></span><br><span data-ttu-id="e462f-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-230">
        - TextBindings</span></span><br><span data-ttu-id="e462f-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-232">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="e462f-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="e462f-233">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-233">- TaskPane</span></span><br><span data-ttu-id="e462f-234">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="e462f-234">
        - Content</span></span><br><span data-ttu-id="e462f-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e462f-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e462f-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e462f-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e462f-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e462f-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e462f-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e462f-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e462f-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e462f-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e462f-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e462f-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e462f-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e462f-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e462f-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e462f-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-245">- BindingEvents</span></span><br><span data-ttu-id="e462f-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-246">
        - CompressedFile</span></span><br><span data-ttu-id="e462f-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-247">
        - DocumentEvents</span></span><br><span data-ttu-id="e462f-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="e462f-248">
        - File</span></span><br><span data-ttu-id="e462f-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-249">
        - ImageCoercion</span></span><br><span data-ttu-id="e462f-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-250">
        - MatrixBindings</span></span><br><span data-ttu-id="e462f-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="e462f-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-252">
        - PdfFile</span></span><br><span data-ttu-id="e462f-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-253">
        - Selection</span></span><br><span data-ttu-id="e462f-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-254">
        - Settings</span></span><br><span data-ttu-id="e462f-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-255">
        - TableBindings</span></span><br><span data-ttu-id="e462f-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-256">
        - TableCoercion</span></span><br><span data-ttu-id="e462f-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-257">
        - TextBindings</span></span><br><span data-ttu-id="e462f-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-259">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="e462f-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="e462f-260">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-260">- TaskPane</span></span><br><span data-ttu-id="e462f-261">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="e462f-261">
        - Content</span></span><br><span data-ttu-id="e462f-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e462f-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e462f-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e462f-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e462f-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e462f-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e462f-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e462f-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e462f-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e462f-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e462f-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e462f-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e462f-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e462f-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e462f-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e462f-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-272">- BindingEvents</span></span><br><span data-ttu-id="e462f-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-273">
        - CompressedFile</span></span><br><span data-ttu-id="e462f-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-274">
        - DocumentEvents</span></span><br><span data-ttu-id="e462f-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="e462f-275">
        - File</span></span><br><span data-ttu-id="e462f-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-276">
        - ImageCoercion</span></span><br><span data-ttu-id="e462f-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-277">
        - MatrixBindings</span></span><br><span data-ttu-id="e462f-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="e462f-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-279">
        - PdfFile</span></span><br><span data-ttu-id="e462f-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-280">
        - Selection</span></span><br><span data-ttu-id="e462f-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-281">
        - Settings</span></span><br><span data-ttu-id="e462f-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-282">
        - TableBindings</span></span><br><span data-ttu-id="e462f-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-283">
        - TableCoercion</span></span><br><span data-ttu-id="e462f-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-284">
        - TextBindings</span></span><br><span data-ttu-id="e462f-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="e462f-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="e462f-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e462f-287">平台</span><span class="sxs-lookup"><span data-stu-id="e462f-287">Platform</span></span></th>
    <th><span data-ttu-id="e462f-288">扩展点</span><span class="sxs-lookup"><span data-stu-id="e462f-288">Extension points</span></span></th>
    <th><span data-ttu-id="e462f-289">API 要求集</span><span class="sxs-lookup"><span data-stu-id="e462f-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="e462f-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e462f-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="e462f-291">Office Online</span></span></td>
    <td> <span data-ttu-id="e462f-292">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="e462f-292">- Mail Read</span></span><br><span data-ttu-id="e462f-293">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="e462f-293">
      - Mail Compose</span></span><br><span data-ttu-id="e462f-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e462f-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e462f-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e462f-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e462f-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e462f-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e462f-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e462f-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e462f-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e462f-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e462f-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e462f-302">不可用</span><span class="sxs-lookup"><span data-stu-id="e462f-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-303">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="e462f-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e462f-304">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="e462f-304">- Mail Read</span></span><br><span data-ttu-id="e462f-305">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="e462f-305">
      - Mail Compose</span></span><br><span data-ttu-id="e462f-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e462f-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e462f-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e462f-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e462f-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="e462f-311">不可用</span><span class="sxs-lookup"><span data-stu-id="e462f-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-312">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="e462f-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e462f-313">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="e462f-313">- Mail Read</span></span><br><span data-ttu-id="e462f-314">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="e462f-314">
      - Mail Compose</span></span><br><span data-ttu-id="e462f-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e462f-316">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="e462f-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e462f-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e462f-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e462f-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e462f-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e462f-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e462f-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e462f-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e462f-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e462f-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e462f-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e462f-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e462f-324">不可用</span><span class="sxs-lookup"><span data-stu-id="e462f-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-325">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="e462f-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e462f-326">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="e462f-326">- Mail Read</span></span><br><span data-ttu-id="e462f-327">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="e462f-327">
      - Mail Compose</span></span><br><span data-ttu-id="e462f-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e462f-329">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="e462f-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e462f-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e462f-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e462f-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e462f-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e462f-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e462f-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e462f-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e462f-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e462f-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e462f-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e462f-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e462f-337">不可用</span><span class="sxs-lookup"><span data-stu-id="e462f-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-338">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="e462f-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e462f-339">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="e462f-339">- Mail Read</span></span><br><span data-ttu-id="e462f-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e462f-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e462f-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e462f-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e462f-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e462f-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e462f-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e462f-346">不可用</span><span class="sxs-lookup"><span data-stu-id="e462f-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-347">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="e462f-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e462f-348">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="e462f-348">- Mail Read</span></span><br><span data-ttu-id="e462f-349">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="e462f-349">
      - Mail Compose</span></span><br><span data-ttu-id="e462f-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e462f-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e462f-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e462f-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e462f-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e462f-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e462f-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e462f-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e462f-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e462f-357">不可用</span><span class="sxs-lookup"><span data-stu-id="e462f-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-358">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="e462f-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="e462f-359">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="e462f-359">- Mail Read</span></span><br><span data-ttu-id="e462f-360">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="e462f-360">
      - Mail Compose</span></span><br><span data-ttu-id="e462f-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e462f-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e462f-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e462f-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e462f-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e462f-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e462f-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e462f-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e462f-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e462f-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e462f-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e462f-369">不可用</span><span class="sxs-lookup"><span data-stu-id="e462f-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-370">Office for Android</span><span class="sxs-lookup"><span data-stu-id="e462f-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="e462f-371">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="e462f-371">- Mail Read</span></span><br><span data-ttu-id="e462f-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e462f-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e462f-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e462f-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e462f-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e462f-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e462f-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e462f-378">不可用</span><span class="sxs-lookup"><span data-stu-id="e462f-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="e462f-379">Word</span><span class="sxs-lookup"><span data-stu-id="e462f-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e462f-380">平台</span><span class="sxs-lookup"><span data-stu-id="e462f-380">Platform</span></span></th>
    <th><span data-ttu-id="e462f-381">扩展点</span><span class="sxs-lookup"><span data-stu-id="e462f-381">Extension points</span></span></th>
    <th><span data-ttu-id="e462f-382">API 要求集</span><span class="sxs-lookup"><span data-stu-id="e462f-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="e462f-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e462f-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="e462f-384">Office Online</span></span></td>
    <td> <span data-ttu-id="e462f-385">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-385">- TaskPane</span></span><br><span data-ttu-id="e462f-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e462f-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e462f-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e462f-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e462f-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-391">- BindingEvents</span></span><br><span data-ttu-id="e462f-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e462f-392">
         - CustomXmlParts</span></span><br><span data-ttu-id="e462f-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-393">
         - DocumentEvents</span></span><br><span data-ttu-id="e462f-394">
         - File</span><span class="sxs-lookup"><span data-stu-id="e462f-394">
         - File</span></span><br><span data-ttu-id="e462f-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-395">
         - HtmlCoercion</span></span><br><span data-ttu-id="e462f-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-396">
         - ImageCoercion</span></span><br><span data-ttu-id="e462f-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-397">
         - MatrixBindings</span></span><br><span data-ttu-id="e462f-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-398">
         - MatrixCoercion</span></span><br><span data-ttu-id="e462f-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-399">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e462f-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-400">
         - PdfFile</span></span><br><span data-ttu-id="e462f-401">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-401">
         - Selection</span></span><br><span data-ttu-id="e462f-402">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-402">
         - Settings</span></span><br><span data-ttu-id="e462f-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-403">
         - TableBindings</span></span><br><span data-ttu-id="e462f-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-404">
         - TableCoercion</span></span><br><span data-ttu-id="e462f-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-405">
         - TextBindings</span></span><br><span data-ttu-id="e462f-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-406">
         - TextCoercion</span></span><br><span data-ttu-id="e462f-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e462f-407">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-408">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="e462f-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e462f-409">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-409">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e462f-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e462f-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-411">- BindingEvents</span></span><br><span data-ttu-id="e462f-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-412">
         - CompressedFile</span></span><br><span data-ttu-id="e462f-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e462f-413">
         - CustomXmlParts</span></span><br><span data-ttu-id="e462f-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-414">
         - DocumentEvents</span></span><br><span data-ttu-id="e462f-415">
         - File</span><span class="sxs-lookup"><span data-stu-id="e462f-415">
         - File</span></span><br><span data-ttu-id="e462f-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-416">
         - HtmlCoercion</span></span><br><span data-ttu-id="e462f-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-417">
         - ImageCoercion</span></span><br><span data-ttu-id="e462f-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-418">
         - MatrixBindings</span></span><br><span data-ttu-id="e462f-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-419">
         - MatrixCoercion</span></span><br><span data-ttu-id="e462f-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-420">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e462f-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-421">
         - PdfFile</span></span><br><span data-ttu-id="e462f-422">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-422">
         - Selection</span></span><br><span data-ttu-id="e462f-423">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-423">
         - Settings</span></span><br><span data-ttu-id="e462f-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-424">
         - TableBindings</span></span><br><span data-ttu-id="e462f-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-425">
         - TableCoercion</span></span><br><span data-ttu-id="e462f-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-426">
         - TextBindings</span></span><br><span data-ttu-id="e462f-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-427">
         - TextCoercion</span></span><br><span data-ttu-id="e462f-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e462f-428">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-429">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="e462f-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e462f-430">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-430">- TaskPane</span></span><br><span data-ttu-id="e462f-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e462f-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e462f-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e462f-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e462f-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-436">- BindingEvents</span></span><br><span data-ttu-id="e462f-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-437">
         - CompressedFile</span></span><br><span data-ttu-id="e462f-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e462f-438">
         - CustomXmlParts</span></span><br><span data-ttu-id="e462f-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-439">
         - DocumentEvents</span></span><br><span data-ttu-id="e462f-440">
         - File</span><span class="sxs-lookup"><span data-stu-id="e462f-440">
         - File</span></span><br><span data-ttu-id="e462f-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-441">
         - HtmlCoercion</span></span><br><span data-ttu-id="e462f-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-442">
         - ImageCoercion</span></span><br><span data-ttu-id="e462f-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-443">
         - MatrixBindings</span></span><br><span data-ttu-id="e462f-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-444">
         - MatrixCoercion</span></span><br><span data-ttu-id="e462f-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-445">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e462f-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-446">
         - PdfFile</span></span><br><span data-ttu-id="e462f-447">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-447">
         - Selection</span></span><br><span data-ttu-id="e462f-448">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-448">
         - Settings</span></span><br><span data-ttu-id="e462f-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-449">
         - TableBindings</span></span><br><span data-ttu-id="e462f-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-450">
         - TableCoercion</span></span><br><span data-ttu-id="e462f-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-451">
         - TextBindings</span></span><br><span data-ttu-id="e462f-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-452">
         - TextCoercion</span></span><br><span data-ttu-id="e462f-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e462f-453">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-454">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="e462f-454">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e462f-455">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-455">- TaskPane</span></span><br><span data-ttu-id="e462f-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e462f-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e462f-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e462f-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e462f-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-461">- BindingEvents</span></span><br><span data-ttu-id="e462f-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-462">
         - CompressedFile</span></span><br><span data-ttu-id="e462f-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e462f-463">
         - CustomXmlParts</span></span><br><span data-ttu-id="e462f-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-464">
         - DocumentEvents</span></span><br><span data-ttu-id="e462f-465">
         - File</span><span class="sxs-lookup"><span data-stu-id="e462f-465">
         - File</span></span><br><span data-ttu-id="e462f-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-466">
         - HtmlCoercion</span></span><br><span data-ttu-id="e462f-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-467">
         - ImageCoercion</span></span><br><span data-ttu-id="e462f-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-468">
         - MatrixBindings</span></span><br><span data-ttu-id="e462f-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-469">
         - MatrixCoercion</span></span><br><span data-ttu-id="e462f-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-470">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e462f-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-471">
         - PdfFile</span></span><br><span data-ttu-id="e462f-472">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-472">
         - Selection</span></span><br><span data-ttu-id="e462f-473">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-473">
         - Settings</span></span><br><span data-ttu-id="e462f-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-474">
         - TableBindings</span></span><br><span data-ttu-id="e462f-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-475">
         - TableCoercion</span></span><br><span data-ttu-id="e462f-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-476">
         - TextBindings</span></span><br><span data-ttu-id="e462f-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-477">
         - TextCoercion</span></span><br><span data-ttu-id="e462f-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e462f-478">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-479">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="e462f-479">Office for iPad</span></span></td>
    <td> <span data-ttu-id="e462f-480">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-480">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e462f-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e462f-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e462f-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e462f-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e462f-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e462f-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-485">- BindingEvents</span></span><br><span data-ttu-id="e462f-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-486">
         - CompressedFile</span></span><br><span data-ttu-id="e462f-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e462f-487">
         - CustomXmlParts</span></span><br><span data-ttu-id="e462f-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-488">
         - DocumentEvents</span></span><br><span data-ttu-id="e462f-489">
         - File</span><span class="sxs-lookup"><span data-stu-id="e462f-489">
         - File</span></span><br><span data-ttu-id="e462f-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-490">
         - HtmlCoercion</span></span><br><span data-ttu-id="e462f-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-491">
         - ImageCoercion</span></span><br><span data-ttu-id="e462f-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-492">
         - MatrixBindings</span></span><br><span data-ttu-id="e462f-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-493">
         - MatrixCoercion</span></span><br><span data-ttu-id="e462f-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-494">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e462f-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-495">
         - PdfFile</span></span><br><span data-ttu-id="e462f-496">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-496">
         - Selection</span></span><br><span data-ttu-id="e462f-497">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-497">
         - Settings</span></span><br><span data-ttu-id="e462f-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-498">
         - TableBindings</span></span><br><span data-ttu-id="e462f-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-499">
         - TableCoercion</span></span><br><span data-ttu-id="e462f-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-500">
         - TextBindings</span></span><br><span data-ttu-id="e462f-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-501">
         - TextCoercion</span></span><br><span data-ttu-id="e462f-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e462f-502">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-503">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="e462f-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e462f-504">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-504">- TaskPane</span></span><br><span data-ttu-id="e462f-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e462f-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e462f-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e462f-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e462f-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e462f-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-510">- BindingEvents</span></span><br><span data-ttu-id="e462f-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-511">
         - CompressedFile</span></span><br><span data-ttu-id="e462f-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e462f-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="e462f-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-513">
         - DocumentEvents</span></span><br><span data-ttu-id="e462f-514">
         - File</span><span class="sxs-lookup"><span data-stu-id="e462f-514">
         - File</span></span><br><span data-ttu-id="e462f-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="e462f-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-516">
         - ImageCoercion</span></span><br><span data-ttu-id="e462f-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-517">
         - MatrixBindings</span></span><br><span data-ttu-id="e462f-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-518">
         - MatrixCoercion</span></span><br><span data-ttu-id="e462f-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-519">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e462f-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-520">
         - PdfFile</span></span><br><span data-ttu-id="e462f-521">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-521">
         - Selection</span></span><br><span data-ttu-id="e462f-522">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-522">
         - Settings</span></span><br><span data-ttu-id="e462f-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-523">
         - TableBindings</span></span><br><span data-ttu-id="e462f-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-524">
         - TableCoercion</span></span><br><span data-ttu-id="e462f-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-525">
         - TextBindings</span></span><br><span data-ttu-id="e462f-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-526">
         - TextCoercion</span></span><br><span data-ttu-id="e462f-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e462f-527">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-528">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="e462f-528">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="e462f-529">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-529">- TaskPane</span></span><br><span data-ttu-id="e462f-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e462f-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e462f-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e462f-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e462f-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e462f-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e462f-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e462f-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-535">- BindingEvents</span></span><br><span data-ttu-id="e462f-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-536">
         - CompressedFile</span></span><br><span data-ttu-id="e462f-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e462f-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="e462f-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-538">
         - DocumentEvents</span></span><br><span data-ttu-id="e462f-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="e462f-539">
         - File</span></span><br><span data-ttu-id="e462f-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="e462f-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-541">
         - ImageCoercion</span></span><br><span data-ttu-id="e462f-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-542">
         - MatrixBindings</span></span><br><span data-ttu-id="e462f-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-543">
         - MatrixCoercion</span></span><br><span data-ttu-id="e462f-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-544">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e462f-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-545">
         - PdfFile</span></span><br><span data-ttu-id="e462f-546">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-546">
         - Selection</span></span><br><span data-ttu-id="e462f-547">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-547">
         - Settings</span></span><br><span data-ttu-id="e462f-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-548">
         - TableBindings</span></span><br><span data-ttu-id="e462f-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-549">
         - TableCoercion</span></span><br><span data-ttu-id="e462f-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e462f-550">
         - TextBindings</span></span><br><span data-ttu-id="e462f-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-551">
         - TextCoercion</span></span><br><span data-ttu-id="e462f-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e462f-552">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="e462f-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="e462f-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e462f-554">平台</span><span class="sxs-lookup"><span data-stu-id="e462f-554">Platform</span></span></th>
    <th><span data-ttu-id="e462f-555">扩展点</span><span class="sxs-lookup"><span data-stu-id="e462f-555">Extension points</span></span></th>
    <th><span data-ttu-id="e462f-556">API 要求集</span><span class="sxs-lookup"><span data-stu-id="e462f-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="e462f-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e462f-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="e462f-558">Office Online</span></span></td>
    <td> <span data-ttu-id="e462f-559">- 内容</span><span class="sxs-lookup"><span data-stu-id="e462f-559">- Content</span></span><br><span data-ttu-id="e462f-560">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-560">
         - TaskPane</span></span><br><span data-ttu-id="e462f-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e462f-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e462f-563">- ActiveView</span></span><br><span data-ttu-id="e462f-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-564">
         - CompressedFile</span></span><br><span data-ttu-id="e462f-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-565">
         - DocumentEvents</span></span><br><span data-ttu-id="e462f-566">
         - File</span><span class="sxs-lookup"><span data-stu-id="e462f-566">
         - File</span></span><br><span data-ttu-id="e462f-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-567">
         - ImageCoercion</span></span><br><span data-ttu-id="e462f-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-568">
         - PdfFile</span></span><br><span data-ttu-id="e462f-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-569">
         - Selection</span></span><br><span data-ttu-id="e462f-570">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-570">
         - Settings</span></span><br><span data-ttu-id="e462f-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-571">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-572">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="e462f-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e462f-573">- 内容</span><span class="sxs-lookup"><span data-stu-id="e462f-573">- Content</span></span><br><span data-ttu-id="e462f-574">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-574">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="e462f-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e462f-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e462f-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e462f-576">- ActiveView</span></span><br><span data-ttu-id="e462f-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-577">
         - CompressedFile</span></span><br><span data-ttu-id="e462f-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-578">
         - DocumentEvents</span></span><br><span data-ttu-id="e462f-579">
         - File</span><span class="sxs-lookup"><span data-stu-id="e462f-579">
         - File</span></span><br><span data-ttu-id="e462f-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-580">
         - ImageCoercion</span></span><br><span data-ttu-id="e462f-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-581">
         - PdfFile</span></span><br><span data-ttu-id="e462f-582">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-582">
         - Selection</span></span><br><span data-ttu-id="e462f-583">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-583">
         - Settings</span></span><br><span data-ttu-id="e462f-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-584">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-585">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="e462f-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e462f-586">- 内容</span><span class="sxs-lookup"><span data-stu-id="e462f-586">- Content</span></span><br><span data-ttu-id="e462f-587">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-587">
         - TaskPane</span></span><br><span data-ttu-id="e462f-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e462f-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e462f-590">- ActiveView</span></span><br><span data-ttu-id="e462f-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-591">
         - CompressedFile</span></span><br><span data-ttu-id="e462f-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-592">
         - DocumentEvents</span></span><br><span data-ttu-id="e462f-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="e462f-593">
         - File</span></span><br><span data-ttu-id="e462f-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-594">
         - ImageCoercion</span></span><br><span data-ttu-id="e462f-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-595">
         - PdfFile</span></span><br><span data-ttu-id="e462f-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-596">
         - Selection</span></span><br><span data-ttu-id="e462f-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-597">
         - Settings</span></span><br><span data-ttu-id="e462f-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-598">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-599">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="e462f-599">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e462f-600">- 内容</span><span class="sxs-lookup"><span data-stu-id="e462f-600">- Content</span></span><br><span data-ttu-id="e462f-601">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-601">
         - TaskPane</span></span><br><span data-ttu-id="e462f-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e462f-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e462f-604">- ActiveView</span></span><br><span data-ttu-id="e462f-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-605">
         - CompressedFile</span></span><br><span data-ttu-id="e462f-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-606">
         - DocumentEvents</span></span><br><span data-ttu-id="e462f-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="e462f-607">
         - File</span></span><br><span data-ttu-id="e462f-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-608">
         - ImageCoercion</span></span><br><span data-ttu-id="e462f-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-609">
         - PdfFile</span></span><br><span data-ttu-id="e462f-610">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-610">
         - Selection</span></span><br><span data-ttu-id="e462f-611">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-611">
         - Settings</span></span><br><span data-ttu-id="e462f-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-612">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-613">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="e462f-613">Office for iPad</span></span></td>
    <td> <span data-ttu-id="e462f-614">- 内容</span><span class="sxs-lookup"><span data-stu-id="e462f-614">- Content</span></span><br><span data-ttu-id="e462f-615">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-615">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="e462f-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="e462f-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e462f-617">- ActiveView</span></span><br><span data-ttu-id="e462f-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-618">
         - CompressedFile</span></span><br><span data-ttu-id="e462f-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-619">
         - DocumentEvents</span></span><br><span data-ttu-id="e462f-620">
         - File</span><span class="sxs-lookup"><span data-stu-id="e462f-620">
         - File</span></span><br><span data-ttu-id="e462f-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-621">
         - PdfFile</span></span><br><span data-ttu-id="e462f-622">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-622">
         - Selection</span></span><br><span data-ttu-id="e462f-623">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-623">
         - Settings</span></span><br><span data-ttu-id="e462f-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-624">
         - TextCoercion</span></span><br><span data-ttu-id="e462f-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-625">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-626">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="e462f-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e462f-627">- 内容</span><span class="sxs-lookup"><span data-stu-id="e462f-627">- Content</span></span><br><span data-ttu-id="e462f-628">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-628">
         - TaskPane</span></span><br><span data-ttu-id="e462f-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e462f-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e462f-631">- ActiveView</span></span><br><span data-ttu-id="e462f-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-632">
         - CompressedFile</span></span><br><span data-ttu-id="e462f-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-633">
         - DocumentEvents</span></span><br><span data-ttu-id="e462f-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="e462f-634">
         - File</span></span><br><span data-ttu-id="e462f-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-635">
         - ImageCoercion</span></span><br><span data-ttu-id="e462f-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-636">
         - PdfFile</span></span><br><span data-ttu-id="e462f-637">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-637">
         - Selection</span></span><br><span data-ttu-id="e462f-638">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-638">
         - Settings</span></span><br><span data-ttu-id="e462f-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-639">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-640">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="e462f-640">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="e462f-641">- 内容</span><span class="sxs-lookup"><span data-stu-id="e462f-641">- Content</span></span><br><span data-ttu-id="e462f-642">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-642">
         - TaskPane</span></span><br><span data-ttu-id="e462f-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e462f-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e462f-645">- ActiveView</span></span><br><span data-ttu-id="e462f-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e462f-646">
         - CompressedFile</span></span><br><span data-ttu-id="e462f-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-647">
         - DocumentEvents</span></span><br><span data-ttu-id="e462f-648">
         - File</span><span class="sxs-lookup"><span data-stu-id="e462f-648">
         - File</span></span><br><span data-ttu-id="e462f-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-649">
         - ImageCoercion</span></span><br><span data-ttu-id="e462f-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e462f-650">
         - PdfFile</span></span><br><span data-ttu-id="e462f-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-651">
         - Selection</span></span><br><span data-ttu-id="e462f-652">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-652">
         - Settings</span></span><br><span data-ttu-id="e462f-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-653">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="e462f-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="e462f-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e462f-655">平台</span><span class="sxs-lookup"><span data-stu-id="e462f-655">Platform</span></span></th>
    <th><span data-ttu-id="e462f-656">扩展点</span><span class="sxs-lookup"><span data-stu-id="e462f-656">Extension points</span></span></th>
    <th><span data-ttu-id="e462f-657">API 要求集</span><span class="sxs-lookup"><span data-stu-id="e462f-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="e462f-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e462f-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="e462f-659">Office Online</span></span></td>
    <td> <span data-ttu-id="e462f-660">- 内容</span><span class="sxs-lookup"><span data-stu-id="e462f-660">- Content</span></span><br><span data-ttu-id="e462f-661">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-661">
         - TaskPane</span></span><br><span data-ttu-id="e462f-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="e462f-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e462f-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="e462f-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e462f-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e462f-665">- DocumentEvents</span></span><br><span data-ttu-id="e462f-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="e462f-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-667">
         - ImageCoercion</span></span><br><span data-ttu-id="e462f-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e462f-668">
         - Settings</span></span><br><span data-ttu-id="e462f-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-669">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="e462f-670">项目</span><span class="sxs-lookup"><span data-stu-id="e462f-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e462f-671">平台</span><span class="sxs-lookup"><span data-stu-id="e462f-671">Platform</span></span></th>
    <th><span data-ttu-id="e462f-672">扩展点</span><span class="sxs-lookup"><span data-stu-id="e462f-672">Extension points</span></span></th>
    <th><span data-ttu-id="e462f-673">API 要求集</span><span class="sxs-lookup"><span data-stu-id="e462f-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="e462f-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e462f-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-675">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="e462f-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e462f-676">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-676">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e462f-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e462f-678">- Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-678">- Selection</span></span><br><span data-ttu-id="e462f-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-679">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-680">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="e462f-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e462f-681">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-681">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e462f-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e462f-683">- Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-683">- Selection</span></span><br><span data-ttu-id="e462f-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-684">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e462f-685">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="e462f-685">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e462f-686">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="e462f-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e462f-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e462f-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e462f-688">- Selection</span><span class="sxs-lookup"><span data-stu-id="e462f-688">- Selection</span></span><br><span data-ttu-id="e462f-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e462f-689">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="e462f-690">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e462f-690">See also</span></span>

- [<span data-ttu-id="e462f-691">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="e462f-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="e462f-692">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="e462f-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="e462f-693">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="e462f-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="e462f-694">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="e462f-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)

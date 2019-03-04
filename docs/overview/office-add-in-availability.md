---
title: Office 外接程序主机和平台可用性
description: Excel、Word、Outlook、PowerPoint、OneNote 和项目支持的要求集。
ms.date: 02/20/2019
localization_priority: Priority
ms.openlocfilehash: a3e9c508a5bae0e7eb660458835b9242d0602818
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/28/2019
ms.locfileid: "30199611"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="8e8bb-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="8e8bb-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="8e8bb-104">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。</span><span class="sxs-lookup"><span data-stu-id="8e8bb-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="8e8bb-105">下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="8e8bb-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="8e8bb-p102">通过 MSI 安装的 Office 2016 的生成号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="8e8bb-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="8e8bb-108">Excel</span><span class="sxs-lookup"><span data-stu-id="8e8bb-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="8e8bb-109">平台</span><span class="sxs-lookup"><span data-stu-id="8e8bb-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="8e8bb-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="8e8bb-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="8e8bb-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="8e8bb-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="8e8bb-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="8e8bb-113">Office Online</span></span></td>
    <td> <span data-ttu-id="8e8bb-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-114">- TaskPane</span></span><br><span data-ttu-id="8e8bb-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="8e8bb-115">
        - Content</span></span><br><span data-ttu-id="8e8bb-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="8e8bb-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8e8bb-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8e8bb-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8e8bb-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8e8bb-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8e8bb-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8e8bb-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8e8bb-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8e8bb-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8e8bb-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8e8bb-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-126">
        - BindingEvents</span></span><br><span data-ttu-id="8e8bb-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-127">
        - CompressedFile</span></span><br><span data-ttu-id="8e8bb-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-128">
        - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-129">
        - File</span></span><br><span data-ttu-id="8e8bb-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-130">
        - MatrixBindings</span></span><br><span data-ttu-id="8e8bb-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e8bb-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-132">
        - Selection</span></span><br><span data-ttu-id="8e8bb-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-133">
        - Settings</span></span><br><span data-ttu-id="8e8bb-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-134">
        - TableBindings</span></span><br><span data-ttu-id="8e8bb-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-135">
        - TableCoercion</span></span><br><span data-ttu-id="8e8bb-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-136">
        - TextBindings</span></span><br><span data-ttu-id="8e8bb-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-138">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="8e8bb-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="8e8bb-139">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-139">
        - TaskPane</span></span><br><span data-ttu-id="8e8bb-140">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="8e8bb-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="8e8bb-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="8e8bb-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="8e8bb-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-142">
        - BindingEvents</span></span><br><span data-ttu-id="8e8bb-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-143">
        - CompressedFile</span></span><br><span data-ttu-id="8e8bb-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-144">
        - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-145">
        - File</span></span><br><span data-ttu-id="8e8bb-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-146">
        - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-147">
        - MatrixBindings</span></span><br><span data-ttu-id="8e8bb-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e8bb-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-149">
        - Selection</span></span><br><span data-ttu-id="8e8bb-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-150">
        - Settings</span></span><br><span data-ttu-id="8e8bb-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-151">
        - TableBindings</span></span><br><span data-ttu-id="8e8bb-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-152">
        - TableCoercion</span></span><br><span data-ttu-id="8e8bb-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-153">
        - TextBindings</span></span><br><span data-ttu-id="8e8bb-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-155">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="8e8bb-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="8e8bb-156">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-156">- TaskPane</span></span><br><span data-ttu-id="8e8bb-157">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="8e8bb-157">
        - Content</span></span></td>
    <td><span data-ttu-id="8e8bb-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8e8bb-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8e8bb-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="8e8bb-160">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-160">- BindingEvents</span></span><br><span data-ttu-id="8e8bb-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-161">
        - CompressedFile</span></span><br><span data-ttu-id="8e8bb-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-162">
        - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-163">
        - File</span></span><br><span data-ttu-id="8e8bb-164">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-164">
        - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-165">
        - MatrixBindings</span></span><br><span data-ttu-id="8e8bb-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e8bb-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-167">
        - Selection</span></span><br><span data-ttu-id="8e8bb-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-168">
        - Settings</span></span><br><span data-ttu-id="8e8bb-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-169">
        - TableBindings</span></span><br><span data-ttu-id="8e8bb-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-170">
        - TableCoercion</span></span><br><span data-ttu-id="8e8bb-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-171">
        - TextBindings</span></span><br><span data-ttu-id="8e8bb-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-173">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="8e8bb-173">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="8e8bb-174">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-174">- TaskPane</span></span><br><span data-ttu-id="8e8bb-175">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="8e8bb-175">
        - Content</span></span><br><span data-ttu-id="8e8bb-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="8e8bb-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8e8bb-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8e8bb-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8e8bb-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8e8bb-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8e8bb-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8e8bb-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8e8bb-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8e8bb-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8e8bb-186">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-186">- BindingEvents</span></span><br><span data-ttu-id="8e8bb-187">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-187">
        - CompressedFile</span></span><br><span data-ttu-id="8e8bb-188">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-188">
        - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-189">
        - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-189">
        - File</span></span><br><span data-ttu-id="8e8bb-190">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-190">
        - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-191">
        - MatrixBindings</span></span><br><span data-ttu-id="8e8bb-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e8bb-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-193">
        - Selection</span></span><br><span data-ttu-id="8e8bb-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-194">
        - Settings</span></span><br><span data-ttu-id="8e8bb-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-195">
        - TableBindings</span></span><br><span data-ttu-id="8e8bb-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-196">
        - TableCoercion</span></span><br><span data-ttu-id="8e8bb-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-197">
        - TextBindings</span></span><br><span data-ttu-id="8e8bb-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-199">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="8e8bb-199">Office for iPad</span></span></td>
    <td><span data-ttu-id="8e8bb-200">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-200">- TaskPane</span></span><br><span data-ttu-id="8e8bb-201">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="8e8bb-201">
        - Content</span></span></td>
    <td><span data-ttu-id="8e8bb-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8e8bb-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8e8bb-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8e8bb-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8e8bb-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8e8bb-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8e8bb-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8e8bb-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8e8bb-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8e8bb-211">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-211">- BindingEvents</span></span><br><span data-ttu-id="8e8bb-212">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-212">
        - CompressedFile</span></span><br><span data-ttu-id="8e8bb-213">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-213">
        - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-214">
        - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-214">
        - File</span></span><br><span data-ttu-id="8e8bb-215">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-215">
        - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-216">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-216">
        - MatrixBindings</span></span><br><span data-ttu-id="8e8bb-217">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-217">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e8bb-218">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-218">
        - Selection</span></span><br><span data-ttu-id="8e8bb-219">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-219">
        - Settings</span></span><br><span data-ttu-id="8e8bb-220">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-220">
        - TableBindings</span></span><br><span data-ttu-id="8e8bb-221">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-221">
        - TableCoercion</span></span><br><span data-ttu-id="8e8bb-222">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-222">
        - TextBindings</span></span><br><span data-ttu-id="8e8bb-223">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-223">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-224">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="8e8bb-224">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="8e8bb-225">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-225">- TaskPane</span></span><br><span data-ttu-id="8e8bb-226">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="8e8bb-226">
        - Content</span></span></td>
    <td><span data-ttu-id="8e8bb-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8e8bb-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8e8bb-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="8e8bb-229">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-229">- BindingEvents</span></span><br><span data-ttu-id="8e8bb-230">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-230">
        - CompressedFile</span></span><br><span data-ttu-id="8e8bb-231">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-231">
        - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-232">
        - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-232">
        - File</span></span><br><span data-ttu-id="8e8bb-233">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-233">
        - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-234">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-234">
        - MatrixBindings</span></span><br><span data-ttu-id="8e8bb-235">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-235">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e8bb-236">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-236">
        - PdfFile</span></span><br><span data-ttu-id="8e8bb-237">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-237">
        - Selection</span></span><br><span data-ttu-id="8e8bb-238">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-238">
        - Settings</span></span><br><span data-ttu-id="8e8bb-239">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-239">
        - TableBindings</span></span><br><span data-ttu-id="8e8bb-240">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-240">
        - TableCoercion</span></span><br><span data-ttu-id="8e8bb-241">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-241">
        - TextBindings</span></span><br><span data-ttu-id="8e8bb-242">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-242">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-243">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="8e8bb-243">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="8e8bb-244">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-244">- TaskPane</span></span><br><span data-ttu-id="8e8bb-245">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="8e8bb-245">
        - Content</span></span><br><span data-ttu-id="8e8bb-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="8e8bb-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8e8bb-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8e8bb-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8e8bb-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8e8bb-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8e8bb-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8e8bb-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8e8bb-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8e8bb-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8e8bb-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-256">- BindingEvents</span></span><br><span data-ttu-id="8e8bb-257">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-257">
        - CompressedFile</span></span><br><span data-ttu-id="8e8bb-258">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-258">
        - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-259">
        - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-259">
        - File</span></span><br><span data-ttu-id="8e8bb-260">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-260">
        - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-261">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-261">
        - MatrixBindings</span></span><br><span data-ttu-id="8e8bb-262">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-262">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e8bb-263">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-263">
        - PdfFile</span></span><br><span data-ttu-id="8e8bb-264">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-264">
        - Selection</span></span><br><span data-ttu-id="8e8bb-265">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-265">
        - Settings</span></span><br><span data-ttu-id="8e8bb-266">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-266">
        - TableBindings</span></span><br><span data-ttu-id="8e8bb-267">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-267">
        - TableCoercion</span></span><br><span data-ttu-id="8e8bb-268">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-268">
        - TextBindings</span></span><br><span data-ttu-id="8e8bb-269">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-269">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="8e8bb-270">Outlook</span><span class="sxs-lookup"><span data-stu-id="8e8bb-270">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8e8bb-271">平台</span><span class="sxs-lookup"><span data-stu-id="8e8bb-271">Platform</span></span></th>
    <th><span data-ttu-id="8e8bb-272">扩展点</span><span class="sxs-lookup"><span data-stu-id="8e8bb-272">Extension points</span></span></th>
    <th><span data-ttu-id="8e8bb-273">API 要求集</span><span class="sxs-lookup"><span data-stu-id="8e8bb-273">API requirement sets</span></span></th>
    <th><span data-ttu-id="8e8bb-274"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-274"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-275">Office Online</span><span class="sxs-lookup"><span data-stu-id="8e8bb-275">Office Online</span></span></td>
    <td> <span data-ttu-id="8e8bb-276">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="8e8bb-276">- Mail Read</span></span><br><span data-ttu-id="8e8bb-277">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="8e8bb-277">
      - Mail Compose</span></span><br><span data-ttu-id="8e8bb-278">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-278">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-279">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-279">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e8bb-280">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-280">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e8bb-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e8bb-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e8bb-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8e8bb-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="8e8bb-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="8e8bb-286">不可用</span><span class="sxs-lookup"><span data-stu-id="8e8bb-286">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-287">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="8e8bb-287">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="8e8bb-288">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="8e8bb-288">- Mail Read</span></span><br><span data-ttu-id="8e8bb-289">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="8e8bb-289">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="8e8bb-290">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-290">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e8bb-291">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-291">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e8bb-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e8bb-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="8e8bb-294">不可用</span><span class="sxs-lookup"><span data-stu-id="8e8bb-294">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-295">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="8e8bb-295">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="8e8bb-296">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="8e8bb-296">- Mail Read</span></span><br><span data-ttu-id="8e8bb-297">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="8e8bb-297">
      - Mail Compose</span></span><br><span data-ttu-id="8e8bb-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="8e8bb-299">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="8e8bb-299">
      - Modules</span></span></td>
    <td> <span data-ttu-id="8e8bb-300">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-300">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e8bb-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e8bb-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e8bb-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e8bb-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8e8bb-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="8e8bb-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="8e8bb-307">不可用</span><span class="sxs-lookup"><span data-stu-id="8e8bb-307">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-308">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="8e8bb-308">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="8e8bb-309">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="8e8bb-309">- Mail Read</span></span><br><span data-ttu-id="8e8bb-310">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="8e8bb-310">
      - Mail Compose</span></span><br><span data-ttu-id="8e8bb-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="8e8bb-312">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="8e8bb-312">
      - Modules</span></span></td>
    <td> <span data-ttu-id="8e8bb-313">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-313">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e8bb-314">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-314">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e8bb-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e8bb-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e8bb-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8e8bb-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="8e8bb-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="8e8bb-320">不可用</span><span class="sxs-lookup"><span data-stu-id="8e8bb-320">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-321">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="8e8bb-321">Office for iOS</span></span></td>
    <td> <span data-ttu-id="8e8bb-322">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="8e8bb-322">- Mail Read</span></span><br><span data-ttu-id="8e8bb-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-324">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-324">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e8bb-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e8bb-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e8bb-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e8bb-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="8e8bb-329">不可用</span><span class="sxs-lookup"><span data-stu-id="8e8bb-329">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-330">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="8e8bb-330">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="8e8bb-331">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="8e8bb-331">- Mail Read</span></span><br><span data-ttu-id="8e8bb-332">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="8e8bb-332">
      - Mail Compose</span></span><br><span data-ttu-id="8e8bb-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-334">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-334">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e8bb-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e8bb-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e8bb-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e8bb-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8e8bb-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="8e8bb-340">不可用</span><span class="sxs-lookup"><span data-stu-id="8e8bb-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-341">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="8e8bb-341">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="8e8bb-342">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="8e8bb-342">- Mail Read</span></span><br><span data-ttu-id="8e8bb-343">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="8e8bb-343">
      - Mail Compose</span></span><br><span data-ttu-id="8e8bb-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-345">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-345">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e8bb-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e8bb-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e8bb-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e8bb-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8e8bb-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="8e8bb-351">不可用</span><span class="sxs-lookup"><span data-stu-id="8e8bb-351">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-352">Office for Android</span><span class="sxs-lookup"><span data-stu-id="8e8bb-352">Office for Android</span></span></td>
    <td> <span data-ttu-id="8e8bb-353">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="8e8bb-353">- Mail Read</span></span><br><span data-ttu-id="8e8bb-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-355">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-355">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e8bb-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e8bb-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e8bb-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e8bb-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="8e8bb-360">不可用</span><span class="sxs-lookup"><span data-stu-id="8e8bb-360">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="8e8bb-361">Word</span><span class="sxs-lookup"><span data-stu-id="8e8bb-361">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8e8bb-362">平台</span><span class="sxs-lookup"><span data-stu-id="8e8bb-362">Platform</span></span></th>
    <th><span data-ttu-id="8e8bb-363">扩展点</span><span class="sxs-lookup"><span data-stu-id="8e8bb-363">Extension points</span></span></th>
    <th><span data-ttu-id="8e8bb-364">API 要求集</span><span class="sxs-lookup"><span data-stu-id="8e8bb-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="8e8bb-365"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-365"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-366">Office Online</span><span class="sxs-lookup"><span data-stu-id="8e8bb-366">Office Online</span></span></td>
    <td> <span data-ttu-id="8e8bb-367">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-367">- TaskPane</span></span><br><span data-ttu-id="8e8bb-368">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-368">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-369">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-369">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8e8bb-370">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-370">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8e8bb-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8e8bb-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-373">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-373">- BindingEvents</span></span><br><span data-ttu-id="8e8bb-374">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e8bb-374">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e8bb-375">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-375">
         - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-376">
         - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-376">
         - File</span></span><br><span data-ttu-id="8e8bb-377">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-377">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e8bb-378">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-378">
         - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-379">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-379">
         - MatrixBindings</span></span><br><span data-ttu-id="8e8bb-380">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-380">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e8bb-381">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-381">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e8bb-382">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-382">
         - PdfFile</span></span><br><span data-ttu-id="8e8bb-383">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-383">
         - Selection</span></span><br><span data-ttu-id="8e8bb-384">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-384">
         - Settings</span></span><br><span data-ttu-id="8e8bb-385">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-385">
         - TableBindings</span></span><br><span data-ttu-id="8e8bb-386">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-386">
         - TableCoercion</span></span><br><span data-ttu-id="8e8bb-387">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-387">
         - TextBindings</span></span><br><span data-ttu-id="8e8bb-388">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-388">
         - TextCoercion</span></span><br><span data-ttu-id="8e8bb-389">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-389">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-390">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="8e8bb-390">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="8e8bb-391">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-391">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8e8bb-392">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="8e8bb-392">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="8e8bb-393">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-393">- BindingEvents</span></span><br><span data-ttu-id="8e8bb-394">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-394">
         - CompressedFile</span></span><br><span data-ttu-id="8e8bb-395">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e8bb-395">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e8bb-396">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-396">
         - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-397">
         - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-397">
         - File</span></span><br><span data-ttu-id="8e8bb-398">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-398">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e8bb-399">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-399">
         - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-400">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-400">
         - MatrixBindings</span></span><br><span data-ttu-id="8e8bb-401">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-401">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e8bb-402">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-402">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e8bb-403">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-403">
         - PdfFile</span></span><br><span data-ttu-id="8e8bb-404">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-404">
         - Selection</span></span><br><span data-ttu-id="8e8bb-405">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-405">
         - Settings</span></span><br><span data-ttu-id="8e8bb-406">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-406">
         - TableBindings</span></span><br><span data-ttu-id="8e8bb-407">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-407">
         - TableCoercion</span></span><br><span data-ttu-id="8e8bb-408">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-408">
         - TextBindings</span></span><br><span data-ttu-id="8e8bb-409">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-409">
         - TextCoercion</span></span><br><span data-ttu-id="8e8bb-410">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-410">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-411">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="8e8bb-411">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="8e8bb-412">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-412">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8e8bb-413">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-413">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8e8bb-414">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8e8bb-414">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="8e8bb-415">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-415">- BindingEvents</span></span><br><span data-ttu-id="8e8bb-416">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-416">
         - CompressedFile</span></span><br><span data-ttu-id="8e8bb-417">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e8bb-417">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e8bb-418">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-418">
         - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-419">
         - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-419">
         - File</span></span><br><span data-ttu-id="8e8bb-420">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-420">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e8bb-421">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-421">
         - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-422">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-422">
         - MatrixBindings</span></span><br><span data-ttu-id="8e8bb-423">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-423">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e8bb-424">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-424">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e8bb-425">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-425">
         - PdfFile</span></span><br><span data-ttu-id="8e8bb-426">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-426">
         - Selection</span></span><br><span data-ttu-id="8e8bb-427">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-427">
         - Settings</span></span><br><span data-ttu-id="8e8bb-428">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-428">
         - TableBindings</span></span><br><span data-ttu-id="8e8bb-429">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-429">
         - TableCoercion</span></span><br><span data-ttu-id="8e8bb-430">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-430">
         - TextBindings</span></span><br><span data-ttu-id="8e8bb-431">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-431">
         - TextCoercion</span></span><br><span data-ttu-id="8e8bb-432">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-432">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-433">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="8e8bb-433">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="8e8bb-434">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-434">- TaskPane</span></span><br><span data-ttu-id="8e8bb-435">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-435">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-436">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-436">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8e8bb-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8e8bb-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8e8bb-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-440">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-440">- BindingEvents</span></span><br><span data-ttu-id="8e8bb-441">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-441">
         - CompressedFile</span></span><br><span data-ttu-id="8e8bb-442">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e8bb-442">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e8bb-443">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-443">
         - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-444">
         - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-444">
         - File</span></span><br><span data-ttu-id="8e8bb-445">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-445">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e8bb-446">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-446">
         - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-447">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-447">
         - MatrixBindings</span></span><br><span data-ttu-id="8e8bb-448">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-448">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e8bb-449">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-449">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e8bb-450">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-450">
         - PdfFile</span></span><br><span data-ttu-id="8e8bb-451">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-451">
         - Selection</span></span><br><span data-ttu-id="8e8bb-452">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-452">
         - Settings</span></span><br><span data-ttu-id="8e8bb-453">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-453">
         - TableBindings</span></span><br><span data-ttu-id="8e8bb-454">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-454">
         - TableCoercion</span></span><br><span data-ttu-id="8e8bb-455">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-455">
         - TextBindings</span></span><br><span data-ttu-id="8e8bb-456">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-456">
         - TextCoercion</span></span><br><span data-ttu-id="8e8bb-457">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-457">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-458">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="8e8bb-458">Office for iPad</span></span></td>
    <td> <span data-ttu-id="8e8bb-459">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-459">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8e8bb-460">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-460">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8e8bb-461">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-461">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8e8bb-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8e8bb-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="8e8bb-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="8e8bb-464">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-464">- BindingEvents</span></span><br><span data-ttu-id="8e8bb-465">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-465">
         - CompressedFile</span></span><br><span data-ttu-id="8e8bb-466">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e8bb-466">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e8bb-467">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-467">
         - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-468">
         - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-468">
         - File</span></span><br><span data-ttu-id="8e8bb-469">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-469">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e8bb-470">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-470">
         - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-471">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-471">
         - MatrixBindings</span></span><br><span data-ttu-id="8e8bb-472">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-472">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e8bb-473">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-473">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e8bb-474">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-474">
         - PdfFile</span></span><br><span data-ttu-id="8e8bb-475">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-475">
         - Selection</span></span><br><span data-ttu-id="8e8bb-476">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-476">
         - Settings</span></span><br><span data-ttu-id="8e8bb-477">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-477">
         - TableBindings</span></span><br><span data-ttu-id="8e8bb-478">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-478">
         - TableCoercion</span></span><br><span data-ttu-id="8e8bb-479">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-479">
         - TextBindings</span></span><br><span data-ttu-id="8e8bb-480">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-480">
         - TextCoercion</span></span><br><span data-ttu-id="8e8bb-481">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-481">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-482">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="8e8bb-482">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="8e8bb-483">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-483">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8e8bb-484">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-484">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8e8bb-485">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8e8bb-485">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="8e8bb-486">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-486">- BindingEvents</span></span><br><span data-ttu-id="8e8bb-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-487">
         - CompressedFile</span></span><br><span data-ttu-id="8e8bb-488">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e8bb-488">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e8bb-489">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-489">
         - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-490">
         - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-490">
         - File</span></span><br><span data-ttu-id="8e8bb-491">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-491">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e8bb-492">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-492">
         - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-493">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-493">
         - MatrixBindings</span></span><br><span data-ttu-id="8e8bb-494">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-494">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e8bb-495">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-495">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e8bb-496">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-496">
         - PdfFile</span></span><br><span data-ttu-id="8e8bb-497">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-497">
         - Selection</span></span><br><span data-ttu-id="8e8bb-498">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-498">
         - Settings</span></span><br><span data-ttu-id="8e8bb-499">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-499">
         - TableBindings</span></span><br><span data-ttu-id="8e8bb-500">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-500">
         - TableCoercion</span></span><br><span data-ttu-id="8e8bb-501">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-501">
         - TextBindings</span></span><br><span data-ttu-id="8e8bb-502">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-502">
         - TextCoercion</span></span><br><span data-ttu-id="8e8bb-503">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-503">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-504">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="8e8bb-504">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="8e8bb-505">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-505">- TaskPane</span></span><br><span data-ttu-id="8e8bb-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8e8bb-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8e8bb-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8e8bb-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="8e8bb-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="8e8bb-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-511">- BindingEvents</span></span><br><span data-ttu-id="8e8bb-512">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-512">
         - CompressedFile</span></span><br><span data-ttu-id="8e8bb-513">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e8bb-513">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e8bb-514">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-514">
         - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-515">
         - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-515">
         - File</span></span><br><span data-ttu-id="8e8bb-516">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-516">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e8bb-517">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-517">
         - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-518">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-518">
         - MatrixBindings</span></span><br><span data-ttu-id="8e8bb-519">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-519">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e8bb-520">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-520">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e8bb-521">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-521">
         - PdfFile</span></span><br><span data-ttu-id="8e8bb-522">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-522">
         - Selection</span></span><br><span data-ttu-id="8e8bb-523">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-523">
         - Settings</span></span><br><span data-ttu-id="8e8bb-524">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-524">
         - TableBindings</span></span><br><span data-ttu-id="8e8bb-525">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-525">
         - TableCoercion</span></span><br><span data-ttu-id="8e8bb-526">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-526">
         - TextBindings</span></span><br><span data-ttu-id="8e8bb-527">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-527">
         - TextCoercion</span></span><br><span data-ttu-id="8e8bb-528">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-528">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="8e8bb-529">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="8e8bb-529">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8e8bb-530">平台</span><span class="sxs-lookup"><span data-stu-id="8e8bb-530">Platform</span></span></th>
    <th><span data-ttu-id="8e8bb-531">扩展点</span><span class="sxs-lookup"><span data-stu-id="8e8bb-531">Extension points</span></span></th>
    <th><span data-ttu-id="8e8bb-532">API 要求集</span><span class="sxs-lookup"><span data-stu-id="8e8bb-532">API requirement sets</span></span></th>
    <th><span data-ttu-id="8e8bb-533"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-533"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-534">Office Online</span><span class="sxs-lookup"><span data-stu-id="8e8bb-534">Office Online</span></span></td>
    <td> <span data-ttu-id="8e8bb-535">- 内容</span><span class="sxs-lookup"><span data-stu-id="8e8bb-535">- Content</span></span><br><span data-ttu-id="8e8bb-536">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-536">
         - TaskPane</span></span><br><span data-ttu-id="8e8bb-537">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-537">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-538">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-538">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-539">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e8bb-539">- ActiveView</span></span><br><span data-ttu-id="8e8bb-540">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-540">
         - CompressedFile</span></span><br><span data-ttu-id="8e8bb-541">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-541">
         - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-542">
         - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-542">
         - File</span></span><br><span data-ttu-id="8e8bb-543">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-543">
         - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-544">
         - PdfFile</span></span><br><span data-ttu-id="8e8bb-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-545">
         - Selection</span></span><br><span data-ttu-id="8e8bb-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-546">
         - Settings</span></span><br><span data-ttu-id="8e8bb-547">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-547">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-548">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="8e8bb-548">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="8e8bb-549">- 内容</span><span class="sxs-lookup"><span data-stu-id="8e8bb-549">- Content</span></span><br><span data-ttu-id="8e8bb-550">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-550">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="8e8bb-551">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="8e8bb-551">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="8e8bb-552">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e8bb-552">- ActiveView</span></span><br><span data-ttu-id="8e8bb-553">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-553">
         - CompressedFile</span></span><br><span data-ttu-id="8e8bb-554">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-554">
         - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-555">
         - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-555">
         - File</span></span><br><span data-ttu-id="8e8bb-556">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-556">
         - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-557">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-557">
         - PdfFile</span></span><br><span data-ttu-id="8e8bb-558">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-558">
         - Selection</span></span><br><span data-ttu-id="8e8bb-559">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-559">
         - Settings</span></span><br><span data-ttu-id="8e8bb-560">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-560">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-561">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="8e8bb-561">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="8e8bb-562">- 内容</span><span class="sxs-lookup"><span data-stu-id="8e8bb-562">- Content</span></span><br><span data-ttu-id="8e8bb-563">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-563">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="8e8bb-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="8e8bb-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="8e8bb-565">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e8bb-565">- ActiveView</span></span><br><span data-ttu-id="8e8bb-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-566">
         - CompressedFile</span></span><br><span data-ttu-id="8e8bb-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-567">
         - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-568">
         - File</span></span><br><span data-ttu-id="8e8bb-569">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-569">
         - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-570">
         - PdfFile</span></span><br><span data-ttu-id="8e8bb-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-571">
         - Selection</span></span><br><span data-ttu-id="8e8bb-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-572">
         - Settings</span></span><br><span data-ttu-id="8e8bb-573">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-573">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-574">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="8e8bb-574">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="8e8bb-575">- 内容</span><span class="sxs-lookup"><span data-stu-id="8e8bb-575">- Content</span></span><br><span data-ttu-id="8e8bb-576">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-576">
         - TaskPane</span></span><br><span data-ttu-id="8e8bb-577">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-577">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-578">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-578">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-579">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e8bb-579">- ActiveView</span></span><br><span data-ttu-id="8e8bb-580">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-580">
         - CompressedFile</span></span><br><span data-ttu-id="8e8bb-581">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-581">
         - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-582">
         - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-582">
         - File</span></span><br><span data-ttu-id="8e8bb-583">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-583">
         - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-584">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-584">
         - PdfFile</span></span><br><span data-ttu-id="8e8bb-585">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-585">
         - Selection</span></span><br><span data-ttu-id="8e8bb-586">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-586">
         - Settings</span></span><br><span data-ttu-id="8e8bb-587">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-587">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-588">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="8e8bb-588">Office for iPad</span></span></td>
    <td> <span data-ttu-id="8e8bb-589">- 内容</span><span class="sxs-lookup"><span data-stu-id="8e8bb-589">- Content</span></span><br><span data-ttu-id="8e8bb-590">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-590">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="8e8bb-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="8e8bb-592">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e8bb-592">- ActiveView</span></span><br><span data-ttu-id="8e8bb-593">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-593">
         - CompressedFile</span></span><br><span data-ttu-id="8e8bb-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-594">
         - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-595">
         - File</span></span><br><span data-ttu-id="8e8bb-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-596">
         - PdfFile</span></span><br><span data-ttu-id="8e8bb-597">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-597">
         - Selection</span></span><br><span data-ttu-id="8e8bb-598">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-598">
         - Settings</span></span><br><span data-ttu-id="8e8bb-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-599">
         - TextCoercion</span></span><br><span data-ttu-id="8e8bb-600">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-600">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-601">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="8e8bb-601">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="8e8bb-602">- 内容</span><span class="sxs-lookup"><span data-stu-id="8e8bb-602">- Content</span></span><br><span data-ttu-id="8e8bb-603">
         - 任务窗格/td></span><span class="sxs-lookup"><span data-stu-id="8e8bb-603">
         - TaskPane/td></span></span> <td> <span data-ttu-id="8e8bb-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="8e8bb-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="8e8bb-605">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e8bb-605">- ActiveView</span></span><br><span data-ttu-id="8e8bb-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-606">
         - CompressedFile</span></span><br><span data-ttu-id="8e8bb-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-607">
         - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-608">
         - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-608">
         - File</span></span><br><span data-ttu-id="8e8bb-609">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-609">
         - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-610">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-610">
         - PdfFile</span></span><br><span data-ttu-id="8e8bb-611">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-611">
         - Selection</span></span><br><span data-ttu-id="8e8bb-612">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-612">
         - Settings</span></span><br><span data-ttu-id="8e8bb-613">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-613">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-614">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="8e8bb-614">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="8e8bb-615">- 内容</span><span class="sxs-lookup"><span data-stu-id="8e8bb-615">- Content</span></span><br><span data-ttu-id="8e8bb-616">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-616">
         - TaskPane</span></span><br><span data-ttu-id="8e8bb-617">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-617">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-619">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e8bb-619">- ActiveView</span></span><br><span data-ttu-id="8e8bb-620">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-620">
         - CompressedFile</span></span><br><span data-ttu-id="8e8bb-621">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-621">
         - DocumentEvents</span></span><br><span data-ttu-id="8e8bb-622">
         - File</span><span class="sxs-lookup"><span data-stu-id="8e8bb-622">
         - File</span></span><br><span data-ttu-id="8e8bb-623">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-623">
         - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e8bb-624">
         - PdfFile</span></span><br><span data-ttu-id="8e8bb-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-625">
         - Selection</span></span><br><span data-ttu-id="8e8bb-626">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-626">
         - Settings</span></span><br><span data-ttu-id="8e8bb-627">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-627">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="8e8bb-628">OneNote</span><span class="sxs-lookup"><span data-stu-id="8e8bb-628">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8e8bb-629">平台</span><span class="sxs-lookup"><span data-stu-id="8e8bb-629">Platform</span></span></th>
    <th><span data-ttu-id="8e8bb-630">扩展点</span><span class="sxs-lookup"><span data-stu-id="8e8bb-630">Extension points</span></span></th>
    <th><span data-ttu-id="8e8bb-631">API 要求集</span><span class="sxs-lookup"><span data-stu-id="8e8bb-631">API requirement sets</span></span></th>
    <th><span data-ttu-id="8e8bb-632"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-632"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-633">Office Online</span><span class="sxs-lookup"><span data-stu-id="8e8bb-633">Office Online</span></span></td>
    <td> <span data-ttu-id="8e8bb-634">- 内容</span><span class="sxs-lookup"><span data-stu-id="8e8bb-634">- Content</span></span><br><span data-ttu-id="8e8bb-635">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-635">
         - TaskPane</span></span><br><span data-ttu-id="8e8bb-636">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-636">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-637">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-637">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="8e8bb-638">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-638">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-639">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e8bb-639">- DocumentEvents</span></span><br><span data-ttu-id="8e8bb-640">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-640">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e8bb-641">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-641">
         - ImageCoercion</span></span><br><span data-ttu-id="8e8bb-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8e8bb-642">
         - Settings</span></span><br><span data-ttu-id="8e8bb-643">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-643">
         - TextCoercion</span></span></td>
  </tr>
</table><span data-ttu-id="8e8bb-644">
\*&ast; - 已添加发布后更新。*


</span><span class="sxs-lookup"><span data-stu-id="8e8bb-644">
\*&ast; - Added with post-release updates.*

</span></span><br/>

## <a name="project"></a><span data-ttu-id="8e8bb-645">项目</span><span class="sxs-lookup"><span data-stu-id="8e8bb-645">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8e8bb-646">平台</span><span class="sxs-lookup"><span data-stu-id="8e8bb-646">Platform</span></span></th>
    <th><span data-ttu-id="8e8bb-647">扩展点</span><span class="sxs-lookup"><span data-stu-id="8e8bb-647">Extension points</span></span></th>
    <th><span data-ttu-id="8e8bb-648">API 要求集</span><span class="sxs-lookup"><span data-stu-id="8e8bb-648">API requirement sets</span></span></th>
    <th><span data-ttu-id="8e8bb-649"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-649"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-650">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="8e8bb-650">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="8e8bb-651">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-651">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8e8bb-652">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-652">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-653">- Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-653">- Selection</span></span><br><span data-ttu-id="8e8bb-654">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-654">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-655">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="8e8bb-655">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="8e8bb-656">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-656">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8e8bb-657">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-657">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-658">- Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-658">- Selection</span></span><br><span data-ttu-id="8e8bb-659">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-659">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e8bb-660">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="8e8bb-660">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="8e8bb-661">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="8e8bb-661">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8e8bb-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e8bb-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e8bb-663">- Selection</span><span class="sxs-lookup"><span data-stu-id="8e8bb-663">- Selection</span></span><br><span data-ttu-id="8e8bb-664">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e8bb-664">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="8e8bb-665">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8e8bb-665">See also</span></span>

- [<span data-ttu-id="8e8bb-666">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="8e8bb-666">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="8e8bb-667">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="8e8bb-667">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="8e8bb-668">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="8e8bb-668">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="8e8bb-669">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="8e8bb-669">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)

---
title: Office 外接程序主机和平台可用性
description: Excel、Word、Outlook、PowerPoint、OneNote 和项目支持的要求集。
ms.date: 11/07/2018
ms.openlocfilehash: c601eac5ed3fcad76b63fff5ae6eeadb7662c8b7
ms.sourcegitcommit: 0adc31ceaba92cb15dc6430c00fe7a96c107c9de
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/09/2018
ms.locfileid: "27210103"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="d4b61-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="d4b61-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="d4b61-104">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。</span><span class="sxs-lookup"><span data-stu-id="d4b61-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="d4b61-105">下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="d4b61-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="d4b61-p102">通过 MSI 安装的 Office 2016 的生成号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="d4b61-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="d4b61-108">Excel</span><span class="sxs-lookup"><span data-stu-id="d4b61-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="d4b61-109">平台</span><span class="sxs-lookup"><span data-stu-id="d4b61-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="d4b61-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="d4b61-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="d4b61-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="d4b61-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="d4b61-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="d4b61-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="d4b61-113">Office Online</span></span></td>
    <td> <span data-ttu-id="d4b61-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-114">- TaskPane</span></span><br><span data-ttu-id="d4b61-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d4b61-115">
        - Content</span></span><br><span data-ttu-id="d4b61-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="d4b61-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="d4b61-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d4b61-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d4b61-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d4b61-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d4b61-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d4b61-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d4b61-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d4b61-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d4b61-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d4b61-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-126">
        - BindingEvents</span></span><br><span data-ttu-id="d4b61-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-127">
        - CompressedFile</span></span><br><span data-ttu-id="d4b61-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-128">
        - DocumentEvents</span></span><br><span data-ttu-id="d4b61-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-129">
        - File</span></span><br><span data-ttu-id="d4b61-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-130">
        - MatrixBindings</span></span><br><span data-ttu-id="d4b61-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="d4b61-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-132">
        - Selection</span></span><br><span data-ttu-id="d4b61-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-133">
        - Settings</span></span><br><span data-ttu-id="d4b61-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-134">
        - TableBindings</span></span><br><span data-ttu-id="d4b61-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-135">
        - TableCoercion</span></span><br><span data-ttu-id="d4b61-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-136">
        - TextBindings</span></span><br><span data-ttu-id="d4b61-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-138">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="d4b61-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="d4b61-139">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-139">
        - TaskPane</span></span><br><span data-ttu-id="d4b61-140">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d4b61-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="d4b61-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d4b61-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-142">
        - BindingEvents</span></span><br><span data-ttu-id="d4b61-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-143">
        - CompressedFile</span></span><br><span data-ttu-id="d4b61-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-144">
        - DocumentEvents</span></span><br><span data-ttu-id="d4b61-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-145">
        - File</span></span><br><span data-ttu-id="d4b61-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-146">
        - ImageCoercion</span></span><br><span data-ttu-id="d4b61-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-147">
        - MatrixBindings</span></span><br><span data-ttu-id="d4b61-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="d4b61-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-149">
        - Selection</span></span><br><span data-ttu-id="d4b61-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-150">
        - Settings</span></span><br><span data-ttu-id="d4b61-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-151">
        - TableBindings</span></span><br><span data-ttu-id="d4b61-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-152">
        - TableCoercion</span></span><br><span data-ttu-id="d4b61-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-153">
        - TextBindings</span></span><br><span data-ttu-id="d4b61-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-155">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="d4b61-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="d4b61-156">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-156">- TaskPane</span></span><br><span data-ttu-id="d4b61-157">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d4b61-157">
        - Content</span></span><br><span data-ttu-id="d4b61-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d4b61-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d4b61-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d4b61-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d4b61-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d4b61-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d4b61-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d4b61-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d4b61-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d4b61-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d4b61-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-168">- BindingEvents</span></span><br><span data-ttu-id="d4b61-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-169">
        - CompressedFile</span></span><br><span data-ttu-id="d4b61-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-170">
        - DocumentEvents</span></span><br><span data-ttu-id="d4b61-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-171">
        - File</span></span><br><span data-ttu-id="d4b61-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-172">
        - ImageCoercion</span></span><br><span data-ttu-id="d4b61-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-173">
        - MatrixBindings</span></span><br><span data-ttu-id="d4b61-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="d4b61-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-175">
        - Selection</span></span><br><span data-ttu-id="d4b61-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-176">
        - Settings</span></span><br><span data-ttu-id="d4b61-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-177">
        - TableBindings</span></span><br><span data-ttu-id="d4b61-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-178">
        - TableCoercion</span></span><br><span data-ttu-id="d4b61-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-179">
        - TextBindings</span></span><br><span data-ttu-id="d4b61-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-181">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="d4b61-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="d4b61-182">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-182">- TaskPane</span></span><br><span data-ttu-id="d4b61-183">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d4b61-183">
        - Content</span></span><br><span data-ttu-id="d4b61-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d4b61-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d4b61-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d4b61-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d4b61-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d4b61-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d4b61-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d4b61-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d4b61-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d4b61-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d4b61-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-194">- BindingEvents</span></span><br><span data-ttu-id="d4b61-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-195">
        - CompressedFile</span></span><br><span data-ttu-id="d4b61-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-196">
        - DocumentEvents</span></span><br><span data-ttu-id="d4b61-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-197">
        - File</span></span><br><span data-ttu-id="d4b61-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-198">
        - ImageCoercion</span></span><br><span data-ttu-id="d4b61-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-199">
        - MatrixBindings</span></span><br><span data-ttu-id="d4b61-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="d4b61-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-201">
        - Selection</span></span><br><span data-ttu-id="d4b61-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-202">
        - Settings</span></span><br><span data-ttu-id="d4b61-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-203">
        - TableBindings</span></span><br><span data-ttu-id="d4b61-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-204">
        - TableCoercion</span></span><br><span data-ttu-id="d4b61-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-205">
        - TextBindings</span></span><br><span data-ttu-id="d4b61-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-207">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="d4b61-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="d4b61-208">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-208">- TaskPane</span></span><br><span data-ttu-id="d4b61-209">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d4b61-209">
        - Content</span></span></td>
    <td><span data-ttu-id="d4b61-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d4b61-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d4b61-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d4b61-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d4b61-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d4b61-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d4b61-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d4b61-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d4b61-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d4b61-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-219">- BindingEvents</span></span><br><span data-ttu-id="d4b61-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-220">
        - CompressedFile</span></span><br><span data-ttu-id="d4b61-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-221">
        - DocumentEvents</span></span><br><span data-ttu-id="d4b61-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-222">
        - File</span></span><br><span data-ttu-id="d4b61-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-223">
        - ImageCoercion</span></span><br><span data-ttu-id="d4b61-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-224">
        - MatrixBindings</span></span><br><span data-ttu-id="d4b61-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="d4b61-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-226">
        - Selection</span></span><br><span data-ttu-id="d4b61-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-227">
        - Settings</span></span><br><span data-ttu-id="d4b61-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-228">
        - TableBindings</span></span><br><span data-ttu-id="d4b61-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-229">
        - TableCoercion</span></span><br><span data-ttu-id="d4b61-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-230">
        - TextBindings</span></span><br><span data-ttu-id="d4b61-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-232">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="d4b61-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="d4b61-233">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-233">- TaskPane</span></span><br><span data-ttu-id="d4b61-234">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d4b61-234">
        - Content</span></span><br><span data-ttu-id="d4b61-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d4b61-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d4b61-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d4b61-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d4b61-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d4b61-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d4b61-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d4b61-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d4b61-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d4b61-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d4b61-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-245">- BindingEvents</span></span><br><span data-ttu-id="d4b61-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-246">
        - CompressedFile</span></span><br><span data-ttu-id="d4b61-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-247">
        - DocumentEvents</span></span><br><span data-ttu-id="d4b61-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-248">
        - File</span></span><br><span data-ttu-id="d4b61-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-249">
        - ImageCoercion</span></span><br><span data-ttu-id="d4b61-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-250">
        - MatrixBindings</span></span><br><span data-ttu-id="d4b61-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="d4b61-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-252">
        - PdfFile</span></span><br><span data-ttu-id="d4b61-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-253">
        - Selection</span></span><br><span data-ttu-id="d4b61-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-254">
        - Settings</span></span><br><span data-ttu-id="d4b61-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-255">
        - TableBindings</span></span><br><span data-ttu-id="d4b61-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-256">
        - TableCoercion</span></span><br><span data-ttu-id="d4b61-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-257">
        - TextBindings</span></span><br><span data-ttu-id="d4b61-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-259">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="d4b61-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="d4b61-260">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-260">- TaskPane</span></span><br><span data-ttu-id="d4b61-261">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="d4b61-261">
        - Content</span></span><br><span data-ttu-id="d4b61-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d4b61-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d4b61-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d4b61-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d4b61-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d4b61-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d4b61-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d4b61-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d4b61-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d4b61-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d4b61-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-272">- BindingEvents</span></span><br><span data-ttu-id="d4b61-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-273">
        - CompressedFile</span></span><br><span data-ttu-id="d4b61-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-274">
        - DocumentEvents</span></span><br><span data-ttu-id="d4b61-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-275">
        - File</span></span><br><span data-ttu-id="d4b61-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-276">
        - ImageCoercion</span></span><br><span data-ttu-id="d4b61-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-277">
        - MatrixBindings</span></span><br><span data-ttu-id="d4b61-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="d4b61-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-279">
        - PdfFile</span></span><br><span data-ttu-id="d4b61-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-280">
        - Selection</span></span><br><span data-ttu-id="d4b61-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-281">
        - Settings</span></span><br><span data-ttu-id="d4b61-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-282">
        - TableBindings</span></span><br><span data-ttu-id="d4b61-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-283">
        - TableCoercion</span></span><br><span data-ttu-id="d4b61-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-284">
        - TextBindings</span></span><br><span data-ttu-id="d4b61-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="d4b61-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="d4b61-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d4b61-287">平台</span><span class="sxs-lookup"><span data-stu-id="d4b61-287">Platform</span></span></th>
    <th><span data-ttu-id="d4b61-288">扩展点</span><span class="sxs-lookup"><span data-stu-id="d4b61-288">Extension points</span></span></th>
    <th><span data-ttu-id="d4b61-289">API 要求集</span><span class="sxs-lookup"><span data-stu-id="d4b61-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="d4b61-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="d4b61-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="d4b61-291">Office Online</span></span></td>
    <td> <span data-ttu-id="d4b61-292">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="d4b61-292">- Mail Read</span></span><br><span data-ttu-id="d4b61-293">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="d4b61-293">
      - Mail Compose</span></span><br><span data-ttu-id="d4b61-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4b61-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4b61-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4b61-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d4b61-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d4b61-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d4b61-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d4b61-302">不可用</span><span class="sxs-lookup"><span data-stu-id="d4b61-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-303">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="d4b61-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="d4b61-304">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="d4b61-304">- Mail Read</span></span><br><span data-ttu-id="d4b61-305">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="d4b61-305">
      - Mail Compose</span></span><br><span data-ttu-id="d4b61-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4b61-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4b61-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4b61-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="d4b61-311">不可用</span><span class="sxs-lookup"><span data-stu-id="d4b61-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-312">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="d4b61-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="d4b61-313">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="d4b61-313">- Mail Read</span></span><br><span data-ttu-id="d4b61-314">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="d4b61-314">
      - Mail Compose</span></span><br><span data-ttu-id="d4b61-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d4b61-316">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="d4b61-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d4b61-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4b61-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4b61-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4b61-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d4b61-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d4b61-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d4b61-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d4b61-324">不可用</span><span class="sxs-lookup"><span data-stu-id="d4b61-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-325">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="d4b61-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="d4b61-326">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="d4b61-326">- Mail Read</span></span><br><span data-ttu-id="d4b61-327">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="d4b61-327">
      - Mail Compose</span></span><br><span data-ttu-id="d4b61-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d4b61-329">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="d4b61-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d4b61-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4b61-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4b61-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4b61-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d4b61-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d4b61-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d4b61-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d4b61-337">不可用</span><span class="sxs-lookup"><span data-stu-id="d4b61-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-338">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="d4b61-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="d4b61-339">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="d4b61-339">- Mail Read</span></span><br><span data-ttu-id="d4b61-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4b61-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4b61-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4b61-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d4b61-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d4b61-346">不可用</span><span class="sxs-lookup"><span data-stu-id="d4b61-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-347">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="d4b61-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="d4b61-348">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="d4b61-348">- Mail Read</span></span><br><span data-ttu-id="d4b61-349">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="d4b61-349">
      - Mail Compose</span></span><br><span data-ttu-id="d4b61-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4b61-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4b61-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4b61-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d4b61-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d4b61-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d4b61-357">不可用</span><span class="sxs-lookup"><span data-stu-id="d4b61-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-358">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="d4b61-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="d4b61-359">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="d4b61-359">- Mail Read</span></span><br><span data-ttu-id="d4b61-360">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="d4b61-360">
      - Mail Compose</span></span><br><span data-ttu-id="d4b61-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4b61-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4b61-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4b61-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d4b61-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d4b61-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d4b61-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d4b61-369">不可用</span><span class="sxs-lookup"><span data-stu-id="d4b61-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-370">Office for Android</span><span class="sxs-lookup"><span data-stu-id="d4b61-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="d4b61-371">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="d4b61-371">- Mail Read</span></span><br><span data-ttu-id="d4b61-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4b61-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4b61-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4b61-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d4b61-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d4b61-378">不可用</span><span class="sxs-lookup"><span data-stu-id="d4b61-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="d4b61-379">Word</span><span class="sxs-lookup"><span data-stu-id="d4b61-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d4b61-380">平台</span><span class="sxs-lookup"><span data-stu-id="d4b61-380">Platform</span></span></th>
    <th><span data-ttu-id="d4b61-381">扩展点</span><span class="sxs-lookup"><span data-stu-id="d4b61-381">Extension points</span></span></th>
    <th><span data-ttu-id="d4b61-382">API 要求集</span><span class="sxs-lookup"><span data-stu-id="d4b61-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="d4b61-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="d4b61-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="d4b61-384">Office Online</span></span></td>
    <td> <span data-ttu-id="d4b61-385">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-385">- TaskPane</span></span><br><span data-ttu-id="d4b61-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d4b61-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d4b61-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d4b61-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4b61-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-391">- BindingEvents</span></span><br><span data-ttu-id="d4b61-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d4b61-392">
         - CustomXmlParts</span></span><br><span data-ttu-id="d4b61-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-393">
         - DocumentEvents</span></span><br><span data-ttu-id="d4b61-394">
         - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-394">
         - File</span></span><br><span data-ttu-id="d4b61-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-395">
         - HtmlCoercion</span></span><br><span data-ttu-id="d4b61-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-396">
         - ImageCoercion</span></span><br><span data-ttu-id="d4b61-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-397">
         - MatrixBindings</span></span><br><span data-ttu-id="d4b61-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-398">
         - MatrixCoercion</span></span><br><span data-ttu-id="d4b61-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-399">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d4b61-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-400">
         - PdfFile</span></span><br><span data-ttu-id="d4b61-401">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-401">
         - Selection</span></span><br><span data-ttu-id="d4b61-402">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-402">
         - Settings</span></span><br><span data-ttu-id="d4b61-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-403">
         - TableBindings</span></span><br><span data-ttu-id="d4b61-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-404">
         - TableCoercion</span></span><br><span data-ttu-id="d4b61-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-405">
         - TextBindings</span></span><br><span data-ttu-id="d4b61-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-406">
         - TextCoercion</span></span><br><span data-ttu-id="d4b61-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-407">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-408">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="d4b61-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="d4b61-409">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-409">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d4b61-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4b61-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-411">- BindingEvents</span></span><br><span data-ttu-id="d4b61-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-412">
         - CompressedFile</span></span><br><span data-ttu-id="d4b61-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d4b61-413">
         - CustomXmlParts</span></span><br><span data-ttu-id="d4b61-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-414">
         - DocumentEvents</span></span><br><span data-ttu-id="d4b61-415">
         - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-415">
         - File</span></span><br><span data-ttu-id="d4b61-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-416">
         - HtmlCoercion</span></span><br><span data-ttu-id="d4b61-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-417">
         - ImageCoercion</span></span><br><span data-ttu-id="d4b61-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-418">
         - MatrixBindings</span></span><br><span data-ttu-id="d4b61-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-419">
         - MatrixCoercion</span></span><br><span data-ttu-id="d4b61-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-420">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d4b61-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-421">
         - PdfFile</span></span><br><span data-ttu-id="d4b61-422">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-422">
         - Selection</span></span><br><span data-ttu-id="d4b61-423">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-423">
         - Settings</span></span><br><span data-ttu-id="d4b61-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-424">
         - TableBindings</span></span><br><span data-ttu-id="d4b61-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-425">
         - TableCoercion</span></span><br><span data-ttu-id="d4b61-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-426">
         - TextBindings</span></span><br><span data-ttu-id="d4b61-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-427">
         - TextCoercion</span></span><br><span data-ttu-id="d4b61-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-428">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-429">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="d4b61-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="d4b61-430">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-430">- TaskPane</span></span><br><span data-ttu-id="d4b61-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d4b61-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d4b61-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d4b61-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4b61-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-436">- BindingEvents</span></span><br><span data-ttu-id="d4b61-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-437">
         - CompressedFile</span></span><br><span data-ttu-id="d4b61-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d4b61-438">
         - CustomXmlParts</span></span><br><span data-ttu-id="d4b61-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-439">
         - DocumentEvents</span></span><br><span data-ttu-id="d4b61-440">
         - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-440">
         - File</span></span><br><span data-ttu-id="d4b61-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-441">
         - HtmlCoercion</span></span><br><span data-ttu-id="d4b61-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-442">
         - ImageCoercion</span></span><br><span data-ttu-id="d4b61-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-443">
         - MatrixBindings</span></span><br><span data-ttu-id="d4b61-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-444">
         - MatrixCoercion</span></span><br><span data-ttu-id="d4b61-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-445">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d4b61-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-446">
         - PdfFile</span></span><br><span data-ttu-id="d4b61-447">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-447">
         - Selection</span></span><br><span data-ttu-id="d4b61-448">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-448">
         - Settings</span></span><br><span data-ttu-id="d4b61-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-449">
         - TableBindings</span></span><br><span data-ttu-id="d4b61-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-450">
         - TableCoercion</span></span><br><span data-ttu-id="d4b61-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-451">
         - TextBindings</span></span><br><span data-ttu-id="d4b61-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-452">
         - TextCoercion</span></span><br><span data-ttu-id="d4b61-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-453">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-454">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="d4b61-454">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="d4b61-455">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-455">- TaskPane</span></span><br><span data-ttu-id="d4b61-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d4b61-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d4b61-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d4b61-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4b61-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-461">- BindingEvents</span></span><br><span data-ttu-id="d4b61-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-462">
         - CompressedFile</span></span><br><span data-ttu-id="d4b61-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d4b61-463">
         - CustomXmlParts</span></span><br><span data-ttu-id="d4b61-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-464">
         - DocumentEvents</span></span><br><span data-ttu-id="d4b61-465">
         - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-465">
         - File</span></span><br><span data-ttu-id="d4b61-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-466">
         - HtmlCoercion</span></span><br><span data-ttu-id="d4b61-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-467">
         - ImageCoercion</span></span><br><span data-ttu-id="d4b61-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-468">
         - MatrixBindings</span></span><br><span data-ttu-id="d4b61-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-469">
         - MatrixCoercion</span></span><br><span data-ttu-id="d4b61-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-470">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d4b61-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-471">
         - PdfFile</span></span><br><span data-ttu-id="d4b61-472">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-472">
         - Selection</span></span><br><span data-ttu-id="d4b61-473">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-473">
         - Settings</span></span><br><span data-ttu-id="d4b61-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-474">
         - TableBindings</span></span><br><span data-ttu-id="d4b61-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-475">
         - TableCoercion</span></span><br><span data-ttu-id="d4b61-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-476">
         - TextBindings</span></span><br><span data-ttu-id="d4b61-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-477">
         - TextCoercion</span></span><br><span data-ttu-id="d4b61-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-478">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-479">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="d4b61-479">Office for iPad</span></span></td>
    <td> <span data-ttu-id="d4b61-480">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-480">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d4b61-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d4b61-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d4b61-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d4b61-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="d4b61-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="d4b61-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-485">- BindingEvents</span></span><br><span data-ttu-id="d4b61-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-486">
         - CompressedFile</span></span><br><span data-ttu-id="d4b61-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d4b61-487">
         - CustomXmlParts</span></span><br><span data-ttu-id="d4b61-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-488">
         - DocumentEvents</span></span><br><span data-ttu-id="d4b61-489">
         - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-489">
         - File</span></span><br><span data-ttu-id="d4b61-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-490">
         - HtmlCoercion</span></span><br><span data-ttu-id="d4b61-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-491">
         - ImageCoercion</span></span><br><span data-ttu-id="d4b61-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-492">
         - MatrixBindings</span></span><br><span data-ttu-id="d4b61-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-493">
         - MatrixCoercion</span></span><br><span data-ttu-id="d4b61-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-494">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d4b61-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-495">
         - PdfFile</span></span><br><span data-ttu-id="d4b61-496">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-496">
         - Selection</span></span><br><span data-ttu-id="d4b61-497">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-497">
         - Settings</span></span><br><span data-ttu-id="d4b61-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-498">
         - TableBindings</span></span><br><span data-ttu-id="d4b61-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-499">
         - TableCoercion</span></span><br><span data-ttu-id="d4b61-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-500">
         - TextBindings</span></span><br><span data-ttu-id="d4b61-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-501">
         - TextCoercion</span></span><br><span data-ttu-id="d4b61-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-502">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-503">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="d4b61-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="d4b61-504">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-504">- TaskPane</span></span><br><span data-ttu-id="d4b61-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d4b61-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d4b61-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d4b61-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="d4b61-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="d4b61-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-510">- BindingEvents</span></span><br><span data-ttu-id="d4b61-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-511">
         - CompressedFile</span></span><br><span data-ttu-id="d4b61-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d4b61-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="d4b61-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-513">
         - DocumentEvents</span></span><br><span data-ttu-id="d4b61-514">
         - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-514">
         - File</span></span><br><span data-ttu-id="d4b61-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="d4b61-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-516">
         - ImageCoercion</span></span><br><span data-ttu-id="d4b61-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-517">
         - MatrixBindings</span></span><br><span data-ttu-id="d4b61-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-518">
         - MatrixCoercion</span></span><br><span data-ttu-id="d4b61-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-519">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d4b61-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-520">
         - PdfFile</span></span><br><span data-ttu-id="d4b61-521">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-521">
         - Selection</span></span><br><span data-ttu-id="d4b61-522">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-522">
         - Settings</span></span><br><span data-ttu-id="d4b61-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-523">
         - TableBindings</span></span><br><span data-ttu-id="d4b61-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-524">
         - TableCoercion</span></span><br><span data-ttu-id="d4b61-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-525">
         - TextBindings</span></span><br><span data-ttu-id="d4b61-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-526">
         - TextCoercion</span></span><br><span data-ttu-id="d4b61-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-527">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-528">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="d4b61-528">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="d4b61-529">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-529">- TaskPane</span></span><br><span data-ttu-id="d4b61-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d4b61-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d4b61-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d4b61-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="d4b61-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="d4b61-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-535">- BindingEvents</span></span><br><span data-ttu-id="d4b61-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-536">
         - CompressedFile</span></span><br><span data-ttu-id="d4b61-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d4b61-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="d4b61-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-538">
         - DocumentEvents</span></span><br><span data-ttu-id="d4b61-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-539">
         - File</span></span><br><span data-ttu-id="d4b61-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="d4b61-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-541">
         - ImageCoercion</span></span><br><span data-ttu-id="d4b61-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-542">
         - MatrixBindings</span></span><br><span data-ttu-id="d4b61-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-543">
         - MatrixCoercion</span></span><br><span data-ttu-id="d4b61-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-544">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d4b61-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-545">
         - PdfFile</span></span><br><span data-ttu-id="d4b61-546">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-546">
         - Selection</span></span><br><span data-ttu-id="d4b61-547">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-547">
         - Settings</span></span><br><span data-ttu-id="d4b61-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-548">
         - TableBindings</span></span><br><span data-ttu-id="d4b61-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-549">
         - TableCoercion</span></span><br><span data-ttu-id="d4b61-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4b61-550">
         - TextBindings</span></span><br><span data-ttu-id="d4b61-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-551">
         - TextCoercion</span></span><br><span data-ttu-id="d4b61-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-552">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="d4b61-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d4b61-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d4b61-554">平台</span><span class="sxs-lookup"><span data-stu-id="d4b61-554">Platform</span></span></th>
    <th><span data-ttu-id="d4b61-555">扩展点</span><span class="sxs-lookup"><span data-stu-id="d4b61-555">Extension points</span></span></th>
    <th><span data-ttu-id="d4b61-556">API 要求集</span><span class="sxs-lookup"><span data-stu-id="d4b61-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="d4b61-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="d4b61-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="d4b61-558">Office Online</span></span></td>
    <td> <span data-ttu-id="d4b61-559">- 内容</span><span class="sxs-lookup"><span data-stu-id="d4b61-559">- Content</span></span><br><span data-ttu-id="d4b61-560">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-560">
         - TaskPane</span></span><br><span data-ttu-id="d4b61-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4b61-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d4b61-563">- ActiveView</span></span><br><span data-ttu-id="d4b61-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-564">
         - CompressedFile</span></span><br><span data-ttu-id="d4b61-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-565">
         - DocumentEvents</span></span><br><span data-ttu-id="d4b61-566">
         - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-566">
         - File</span></span><br><span data-ttu-id="d4b61-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-567">
         - ImageCoercion</span></span><br><span data-ttu-id="d4b61-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-568">
         - PdfFile</span></span><br><span data-ttu-id="d4b61-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-569">
         - Selection</span></span><br><span data-ttu-id="d4b61-570">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-570">
         - Settings</span></span><br><span data-ttu-id="d4b61-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-571">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-572">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="d4b61-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="d4b61-573">- 内容</span><span class="sxs-lookup"><span data-stu-id="d4b61-573">- Content</span></span><br><span data-ttu-id="d4b61-574">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-574">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="d4b61-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="d4b61-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="d4b61-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d4b61-576">- ActiveView</span></span><br><span data-ttu-id="d4b61-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-577">
         - CompressedFile</span></span><br><span data-ttu-id="d4b61-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-578">
         - DocumentEvents</span></span><br><span data-ttu-id="d4b61-579">
         - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-579">
         - File</span></span><br><span data-ttu-id="d4b61-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-580">
         - ImageCoercion</span></span><br><span data-ttu-id="d4b61-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-581">
         - PdfFile</span></span><br><span data-ttu-id="d4b61-582">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-582">
         - Selection</span></span><br><span data-ttu-id="d4b61-583">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-583">
         - Settings</span></span><br><span data-ttu-id="d4b61-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-584">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-585">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="d4b61-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="d4b61-586">- 内容</span><span class="sxs-lookup"><span data-stu-id="d4b61-586">- Content</span></span><br><span data-ttu-id="d4b61-587">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-587">
         - TaskPane</span></span><br><span data-ttu-id="d4b61-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4b61-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d4b61-590">- ActiveView</span></span><br><span data-ttu-id="d4b61-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-591">
         - CompressedFile</span></span><br><span data-ttu-id="d4b61-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-592">
         - DocumentEvents</span></span><br><span data-ttu-id="d4b61-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-593">
         - File</span></span><br><span data-ttu-id="d4b61-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-594">
         - ImageCoercion</span></span><br><span data-ttu-id="d4b61-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-595">
         - PdfFile</span></span><br><span data-ttu-id="d4b61-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-596">
         - Selection</span></span><br><span data-ttu-id="d4b61-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-597">
         - Settings</span></span><br><span data-ttu-id="d4b61-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-598">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-599">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="d4b61-599">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="d4b61-600">- 内容</span><span class="sxs-lookup"><span data-stu-id="d4b61-600">- Content</span></span><br><span data-ttu-id="d4b61-601">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-601">
         - TaskPane</span></span><br><span data-ttu-id="d4b61-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4b61-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d4b61-604">- ActiveView</span></span><br><span data-ttu-id="d4b61-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-605">
         - CompressedFile</span></span><br><span data-ttu-id="d4b61-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-606">
         - DocumentEvents</span></span><br><span data-ttu-id="d4b61-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-607">
         - File</span></span><br><span data-ttu-id="d4b61-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-608">
         - ImageCoercion</span></span><br><span data-ttu-id="d4b61-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-609">
         - PdfFile</span></span><br><span data-ttu-id="d4b61-610">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-610">
         - Selection</span></span><br><span data-ttu-id="d4b61-611">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-611">
         - Settings</span></span><br><span data-ttu-id="d4b61-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-612">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-613">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="d4b61-613">Office for iPad</span></span></td>
    <td> <span data-ttu-id="d4b61-614">- 内容</span><span class="sxs-lookup"><span data-stu-id="d4b61-614">- Content</span></span><br><span data-ttu-id="d4b61-615">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-615">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d4b61-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="d4b61-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d4b61-617">- ActiveView</span></span><br><span data-ttu-id="d4b61-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-618">
         - CompressedFile</span></span><br><span data-ttu-id="d4b61-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-619">
         - DocumentEvents</span></span><br><span data-ttu-id="d4b61-620">
         - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-620">
         - File</span></span><br><span data-ttu-id="d4b61-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-621">
         - PdfFile</span></span><br><span data-ttu-id="d4b61-622">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-622">
         - Selection</span></span><br><span data-ttu-id="d4b61-623">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-623">
         - Settings</span></span><br><span data-ttu-id="d4b61-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-624">
         - TextCoercion</span></span><br><span data-ttu-id="d4b61-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-625">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-626">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="d4b61-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="d4b61-627">- 内容</span><span class="sxs-lookup"><span data-stu-id="d4b61-627">- Content</span></span><br><span data-ttu-id="d4b61-628">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-628">
         - TaskPane</span></span><br><span data-ttu-id="d4b61-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4b61-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d4b61-631">- ActiveView</span></span><br><span data-ttu-id="d4b61-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-632">
         - CompressedFile</span></span><br><span data-ttu-id="d4b61-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-633">
         - DocumentEvents</span></span><br><span data-ttu-id="d4b61-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-634">
         - File</span></span><br><span data-ttu-id="d4b61-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-635">
         - ImageCoercion</span></span><br><span data-ttu-id="d4b61-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-636">
         - PdfFile</span></span><br><span data-ttu-id="d4b61-637">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-637">
         - Selection</span></span><br><span data-ttu-id="d4b61-638">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-638">
         - Settings</span></span><br><span data-ttu-id="d4b61-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-639">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-640">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="d4b61-640">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="d4b61-641">- 内容</span><span class="sxs-lookup"><span data-stu-id="d4b61-641">- Content</span></span><br><span data-ttu-id="d4b61-642">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-642">
         - TaskPane</span></span><br><span data-ttu-id="d4b61-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4b61-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d4b61-645">- ActiveView</span></span><br><span data-ttu-id="d4b61-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-646">
         - CompressedFile</span></span><br><span data-ttu-id="d4b61-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-647">
         - DocumentEvents</span></span><br><span data-ttu-id="d4b61-648">
         - File</span><span class="sxs-lookup"><span data-stu-id="d4b61-648">
         - File</span></span><br><span data-ttu-id="d4b61-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-649">
         - ImageCoercion</span></span><br><span data-ttu-id="d4b61-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4b61-650">
         - PdfFile</span></span><br><span data-ttu-id="d4b61-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-651">
         - Selection</span></span><br><span data-ttu-id="d4b61-652">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-652">
         - Settings</span></span><br><span data-ttu-id="d4b61-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-653">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="d4b61-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="d4b61-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d4b61-655">平台</span><span class="sxs-lookup"><span data-stu-id="d4b61-655">Platform</span></span></th>
    <th><span data-ttu-id="d4b61-656">扩展点</span><span class="sxs-lookup"><span data-stu-id="d4b61-656">Extension points</span></span></th>
    <th><span data-ttu-id="d4b61-657">API 要求集</span><span class="sxs-lookup"><span data-stu-id="d4b61-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="d4b61-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="d4b61-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="d4b61-659">Office Online</span></span></td>
    <td> <span data-ttu-id="d4b61-660">- 内容</span><span class="sxs-lookup"><span data-stu-id="d4b61-660">- Content</span></span><br><span data-ttu-id="d4b61-661">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-661">
         - TaskPane</span></span><br><span data-ttu-id="d4b61-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4b61-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="d4b61-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4b61-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4b61-665">- DocumentEvents</span></span><br><span data-ttu-id="d4b61-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="d4b61-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-667">
         - ImageCoercion</span></span><br><span data-ttu-id="d4b61-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d4b61-668">
         - Settings</span></span><br><span data-ttu-id="d4b61-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-669">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="d4b61-670">项目</span><span class="sxs-lookup"><span data-stu-id="d4b61-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d4b61-671">平台</span><span class="sxs-lookup"><span data-stu-id="d4b61-671">Platform</span></span></th>
    <th><span data-ttu-id="d4b61-672">扩展点</span><span class="sxs-lookup"><span data-stu-id="d4b61-672">Extension points</span></span></th>
    <th><span data-ttu-id="d4b61-673">API 要求集</span><span class="sxs-lookup"><span data-stu-id="d4b61-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="d4b61-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="d4b61-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-675">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="d4b61-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="d4b61-676">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-676">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d4b61-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4b61-678">- Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-678">- Selection</span></span><br><span data-ttu-id="d4b61-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-679">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-680">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="d4b61-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="d4b61-681">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-681">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d4b61-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4b61-683">- Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-683">- Selection</span></span><br><span data-ttu-id="d4b61-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-684">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4b61-685">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="d4b61-685">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="d4b61-686">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="d4b61-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d4b61-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4b61-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4b61-688">- Selection</span><span class="sxs-lookup"><span data-stu-id="d4b61-688">- Selection</span></span><br><span data-ttu-id="d4b61-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4b61-689">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="d4b61-690">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d4b61-690">See also</span></span>

- [<span data-ttu-id="d4b61-691">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="d4b61-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="d4b61-692">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="d4b61-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="d4b61-693">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="d4b61-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="d4b61-694">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="d4b61-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)

---
title: Office 外接程序主机和平台可用性
description: Excel、Word、Outlook、PowerPoint、OneNote 和项目支持的要求集。
ms.date: 11/07/2018
ms.openlocfilehash: 9490fca9663737e2397de159169b545e3900289f
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/28/2018
ms.locfileid: "27458039"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="1e9c1-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="1e9c1-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="1e9c1-104">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。</span><span class="sxs-lookup"><span data-stu-id="1e9c1-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="1e9c1-105">下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="1e9c1-105">The following tables contain the available platforms, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="1e9c1-p102">通过 MSI 安装的 Office 2016 的生成号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="1e9c1-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="1e9c1-108">Excel</span><span class="sxs-lookup"><span data-stu-id="1e9c1-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="1e9c1-109">平台</span><span class="sxs-lookup"><span data-stu-id="1e9c1-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="1e9c1-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="1e9c1-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="1e9c1-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="1e9c1-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="1e9c1-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="1e9c1-113">Office Online</span></span></td>
    <td> <span data-ttu-id="1e9c1-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-114">- TaskPane</span></span><br><span data-ttu-id="1e9c1-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="1e9c1-115">
        - Content</span></span><br><span data-ttu-id="1e9c1-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="1e9c1-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="1e9c1-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1e9c1-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1e9c1-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1e9c1-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1e9c1-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1e9c1-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1e9c1-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1e9c1-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1e9c1-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1e9c1-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-126">
        - BindingEvents</span></span><br><span data-ttu-id="1e9c1-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-127">
        - CompressedFile</span></span><br><span data-ttu-id="1e9c1-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-128">
        - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-129">
        - File</span></span><br><span data-ttu-id="1e9c1-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-130">
        - MatrixBindings</span></span><br><span data-ttu-id="1e9c1-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="1e9c1-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-132">
        - Selection</span></span><br><span data-ttu-id="1e9c1-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-133">
        - Settings</span></span><br><span data-ttu-id="1e9c1-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-134">
        - TableBindings</span></span><br><span data-ttu-id="1e9c1-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-135">
        - TableCoercion</span></span><br><span data-ttu-id="1e9c1-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-136">
        - TextBindings</span></span><br><span data-ttu-id="1e9c1-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-138">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="1e9c1-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="1e9c1-139">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-139">
        - TaskPane</span></span><br><span data-ttu-id="1e9c1-140">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="1e9c1-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="1e9c1-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1e9c1-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-142">
        - BindingEvents</span></span><br><span data-ttu-id="1e9c1-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-143">
        - CompressedFile</span></span><br><span data-ttu-id="1e9c1-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-144">
        - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-145">
        - File</span></span><br><span data-ttu-id="1e9c1-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-146">
        - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-147">
        - MatrixBindings</span></span><br><span data-ttu-id="1e9c1-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="1e9c1-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-149">
        - Selection</span></span><br><span data-ttu-id="1e9c1-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-150">
        - Settings</span></span><br><span data-ttu-id="1e9c1-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-151">
        - TableBindings</span></span><br><span data-ttu-id="1e9c1-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-152">
        - TableCoercion</span></span><br><span data-ttu-id="1e9c1-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-153">
        - TextBindings</span></span><br><span data-ttu-id="1e9c1-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-155">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="1e9c1-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="1e9c1-156">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-156">- TaskPane</span></span><br><span data-ttu-id="1e9c1-157">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="1e9c1-157">
        - Content</span></span><br><span data-ttu-id="1e9c1-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1e9c1-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1e9c1-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1e9c1-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1e9c1-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1e9c1-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1e9c1-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1e9c1-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1e9c1-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1e9c1-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1e9c1-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-168">- BindingEvents</span></span><br><span data-ttu-id="1e9c1-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-169">
        - CompressedFile</span></span><br><span data-ttu-id="1e9c1-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-170">
        - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-171">
        - File</span></span><br><span data-ttu-id="1e9c1-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-172">
        - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-173">
        - MatrixBindings</span></span><br><span data-ttu-id="1e9c1-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="1e9c1-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-175">
        - Selection</span></span><br><span data-ttu-id="1e9c1-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-176">
        - Settings</span></span><br><span data-ttu-id="1e9c1-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-177">
        - TableBindings</span></span><br><span data-ttu-id="1e9c1-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-178">
        - TableCoercion</span></span><br><span data-ttu-id="1e9c1-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-179">
        - TextBindings</span></span><br><span data-ttu-id="1e9c1-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-181">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="1e9c1-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="1e9c1-182">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-182">- TaskPane</span></span><br><span data-ttu-id="1e9c1-183">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="1e9c1-183">
        - Content</span></span><br><span data-ttu-id="1e9c1-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1e9c1-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1e9c1-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1e9c1-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1e9c1-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1e9c1-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1e9c1-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1e9c1-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1e9c1-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1e9c1-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1e9c1-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-194">- BindingEvents</span></span><br><span data-ttu-id="1e9c1-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-195">
        - CompressedFile</span></span><br><span data-ttu-id="1e9c1-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-196">
        - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-197">
        - File</span></span><br><span data-ttu-id="1e9c1-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-198">
        - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-199">
        - MatrixBindings</span></span><br><span data-ttu-id="1e9c1-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="1e9c1-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-201">
        - Selection</span></span><br><span data-ttu-id="1e9c1-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-202">
        - Settings</span></span><br><span data-ttu-id="1e9c1-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-203">
        - TableBindings</span></span><br><span data-ttu-id="1e9c1-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-204">
        - TableCoercion</span></span><br><span data-ttu-id="1e9c1-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-205">
        - TextBindings</span></span><br><span data-ttu-id="1e9c1-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-207">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="1e9c1-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="1e9c1-208">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-208">- TaskPane</span></span><br><span data-ttu-id="1e9c1-209">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="1e9c1-209">
        - Content</span></span></td>
    <td><span data-ttu-id="1e9c1-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1e9c1-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1e9c1-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1e9c1-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1e9c1-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1e9c1-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1e9c1-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1e9c1-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1e9c1-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1e9c1-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-219">- BindingEvents</span></span><br><span data-ttu-id="1e9c1-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-220">
        - CompressedFile</span></span><br><span data-ttu-id="1e9c1-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-221">
        - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-222">
        - File</span></span><br><span data-ttu-id="1e9c1-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-223">
        - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-224">
        - MatrixBindings</span></span><br><span data-ttu-id="1e9c1-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="1e9c1-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-226">
        - Selection</span></span><br><span data-ttu-id="1e9c1-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-227">
        - Settings</span></span><br><span data-ttu-id="1e9c1-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-228">
        - TableBindings</span></span><br><span data-ttu-id="1e9c1-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-229">
        - TableCoercion</span></span><br><span data-ttu-id="1e9c1-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-230">
        - TextBindings</span></span><br><span data-ttu-id="1e9c1-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-232">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="1e9c1-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="1e9c1-233">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-233">- TaskPane</span></span><br><span data-ttu-id="1e9c1-234">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="1e9c1-234">
        - Content</span></span><br><span data-ttu-id="1e9c1-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1e9c1-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1e9c1-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1e9c1-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1e9c1-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1e9c1-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1e9c1-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1e9c1-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1e9c1-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1e9c1-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1e9c1-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-245">- BindingEvents</span></span><br><span data-ttu-id="1e9c1-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-246">
        - CompressedFile</span></span><br><span data-ttu-id="1e9c1-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-247">
        - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-248">
        - File</span></span><br><span data-ttu-id="1e9c1-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-249">
        - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-250">
        - MatrixBindings</span></span><br><span data-ttu-id="1e9c1-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="1e9c1-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-252">
        - PdfFile</span></span><br><span data-ttu-id="1e9c1-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-253">
        - Selection</span></span><br><span data-ttu-id="1e9c1-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-254">
        - Settings</span></span><br><span data-ttu-id="1e9c1-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-255">
        - TableBindings</span></span><br><span data-ttu-id="1e9c1-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-256">
        - TableCoercion</span></span><br><span data-ttu-id="1e9c1-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-257">
        - TextBindings</span></span><br><span data-ttu-id="1e9c1-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-259">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="1e9c1-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="1e9c1-260">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-260">- TaskPane</span></span><br><span data-ttu-id="1e9c1-261">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="1e9c1-261">
        - Content</span></span><br><span data-ttu-id="1e9c1-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1e9c1-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1e9c1-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1e9c1-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1e9c1-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1e9c1-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1e9c1-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1e9c1-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1e9c1-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1e9c1-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1e9c1-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-272">- BindingEvents</span></span><br><span data-ttu-id="1e9c1-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-273">
        - CompressedFile</span></span><br><span data-ttu-id="1e9c1-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-274">
        - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-275">
        - File</span></span><br><span data-ttu-id="1e9c1-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-276">
        - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-277">
        - MatrixBindings</span></span><br><span data-ttu-id="1e9c1-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="1e9c1-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-279">
        - PdfFile</span></span><br><span data-ttu-id="1e9c1-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-280">
        - Selection</span></span><br><span data-ttu-id="1e9c1-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-281">
        - Settings</span></span><br><span data-ttu-id="1e9c1-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-282">
        - TableBindings</span></span><br><span data-ttu-id="1e9c1-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-283">
        - TableCoercion</span></span><br><span data-ttu-id="1e9c1-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-284">
        - TextBindings</span></span><br><span data-ttu-id="1e9c1-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="1e9c1-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="1e9c1-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1e9c1-287">平台</span><span class="sxs-lookup"><span data-stu-id="1e9c1-287">Platform</span></span></th>
    <th><span data-ttu-id="1e9c1-288">扩展点</span><span class="sxs-lookup"><span data-stu-id="1e9c1-288">Extension points</span></span></th>
    <th><span data-ttu-id="1e9c1-289">API 要求集</span><span class="sxs-lookup"><span data-stu-id="1e9c1-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="1e9c1-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="1e9c1-291">Office Online</span></span></td>
    <td> <span data-ttu-id="1e9c1-292">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="1e9c1-292">- Mail Read</span></span><br><span data-ttu-id="1e9c1-293">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="1e9c1-293">
      - Mail Compose</span></span><br><span data-ttu-id="1e9c1-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1e9c1-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1e9c1-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1e9c1-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1e9c1-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1e9c1-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1e9c1-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1e9c1-302">不可用</span><span class="sxs-lookup"><span data-stu-id="1e9c1-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-303">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="1e9c1-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="1e9c1-304">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="1e9c1-304">- Mail Read</span></span><br><span data-ttu-id="1e9c1-305">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="1e9c1-305">
      - Mail Compose</span></span><br><span data-ttu-id="1e9c1-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1e9c1-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1e9c1-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1e9c1-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="1e9c1-311">不可用</span><span class="sxs-lookup"><span data-stu-id="1e9c1-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-312">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="1e9c1-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="1e9c1-313">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="1e9c1-313">- Mail Read</span></span><br><span data-ttu-id="1e9c1-314">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="1e9c1-314">
      - Mail Compose</span></span><br><span data-ttu-id="1e9c1-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1e9c1-316">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="1e9c1-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="1e9c1-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1e9c1-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1e9c1-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1e9c1-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1e9c1-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1e9c1-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1e9c1-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1e9c1-324">不可用</span><span class="sxs-lookup"><span data-stu-id="1e9c1-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-325">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="1e9c1-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="1e9c1-326">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="1e9c1-326">- Mail Read</span></span><br><span data-ttu-id="1e9c1-327">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="1e9c1-327">
      - Mail Compose</span></span><br><span data-ttu-id="1e9c1-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1e9c1-329">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="1e9c1-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="1e9c1-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1e9c1-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1e9c1-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1e9c1-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1e9c1-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1e9c1-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1e9c1-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1e9c1-337">不可用</span><span class="sxs-lookup"><span data-stu-id="1e9c1-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-338">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="1e9c1-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="1e9c1-339">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="1e9c1-339">- Mail Read</span></span><br><span data-ttu-id="1e9c1-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1e9c1-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1e9c1-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1e9c1-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1e9c1-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="1e9c1-346">不可用</span><span class="sxs-lookup"><span data-stu-id="1e9c1-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-347">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="1e9c1-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="1e9c1-348">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="1e9c1-348">- Mail Read</span></span><br><span data-ttu-id="1e9c1-349">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="1e9c1-349">
      - Mail Compose</span></span><br><span data-ttu-id="1e9c1-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1e9c1-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1e9c1-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1e9c1-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1e9c1-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1e9c1-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="1e9c1-357">不可用</span><span class="sxs-lookup"><span data-stu-id="1e9c1-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-358">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="1e9c1-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="1e9c1-359">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="1e9c1-359">- Mail Read</span></span><br><span data-ttu-id="1e9c1-360">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="1e9c1-360">
      - Mail Compose</span></span><br><span data-ttu-id="1e9c1-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1e9c1-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1e9c1-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1e9c1-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1e9c1-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1e9c1-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1e9c1-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1e9c1-369">不可用</span><span class="sxs-lookup"><span data-stu-id="1e9c1-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-370">Office for Android</span><span class="sxs-lookup"><span data-stu-id="1e9c1-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="1e9c1-371">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="1e9c1-371">- Mail Read</span></span><br><span data-ttu-id="1e9c1-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1e9c1-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1e9c1-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1e9c1-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1e9c1-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="1e9c1-378">不可用</span><span class="sxs-lookup"><span data-stu-id="1e9c1-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="1e9c1-379">Word</span><span class="sxs-lookup"><span data-stu-id="1e9c1-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1e9c1-380">平台</span><span class="sxs-lookup"><span data-stu-id="1e9c1-380">Platform</span></span></th>
    <th><span data-ttu-id="1e9c1-381">扩展点</span><span class="sxs-lookup"><span data-stu-id="1e9c1-381">Extension points</span></span></th>
    <th><span data-ttu-id="1e9c1-382">API 要求集</span><span class="sxs-lookup"><span data-stu-id="1e9c1-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="1e9c1-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="1e9c1-384">Office Online</span></span></td>
    <td> <span data-ttu-id="1e9c1-385">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-385">- TaskPane</span></span><br><span data-ttu-id="1e9c1-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1e9c1-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1e9c1-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1e9c1-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-391">- BindingEvents</span></span><br><span data-ttu-id="1e9c1-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1e9c1-392">
         - CustomXmlParts</span></span><br><span data-ttu-id="1e9c1-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-393">
         - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-394">
         - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-394">
         - File</span></span><br><span data-ttu-id="1e9c1-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-395">
         - HtmlCoercion</span></span><br><span data-ttu-id="1e9c1-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-396">
         - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-397">
         - MatrixBindings</span></span><br><span data-ttu-id="1e9c1-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-398">
         - MatrixCoercion</span></span><br><span data-ttu-id="1e9c1-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-399">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1e9c1-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-400">
         - PdfFile</span></span><br><span data-ttu-id="1e9c1-401">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-401">
         - Selection</span></span><br><span data-ttu-id="1e9c1-402">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-402">
         - Settings</span></span><br><span data-ttu-id="1e9c1-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-403">
         - TableBindings</span></span><br><span data-ttu-id="1e9c1-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-404">
         - TableCoercion</span></span><br><span data-ttu-id="1e9c1-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-405">
         - TextBindings</span></span><br><span data-ttu-id="1e9c1-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-406">
         - TextCoercion</span></span><br><span data-ttu-id="1e9c1-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-407">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-408">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="1e9c1-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="1e9c1-409">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-409">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1e9c1-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-411">- BindingEvents</span></span><br><span data-ttu-id="1e9c1-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-412">
         - CompressedFile</span></span><br><span data-ttu-id="1e9c1-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1e9c1-413">
         - CustomXmlParts</span></span><br><span data-ttu-id="1e9c1-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-414">
         - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-415">
         - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-415">
         - File</span></span><br><span data-ttu-id="1e9c1-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-416">
         - HtmlCoercion</span></span><br><span data-ttu-id="1e9c1-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-417">
         - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-418">
         - MatrixBindings</span></span><br><span data-ttu-id="1e9c1-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-419">
         - MatrixCoercion</span></span><br><span data-ttu-id="1e9c1-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-420">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1e9c1-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-421">
         - PdfFile</span></span><br><span data-ttu-id="1e9c1-422">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-422">
         - Selection</span></span><br><span data-ttu-id="1e9c1-423">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-423">
         - Settings</span></span><br><span data-ttu-id="1e9c1-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-424">
         - TableBindings</span></span><br><span data-ttu-id="1e9c1-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-425">
         - TableCoercion</span></span><br><span data-ttu-id="1e9c1-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-426">
         - TextBindings</span></span><br><span data-ttu-id="1e9c1-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-427">
         - TextCoercion</span></span><br><span data-ttu-id="1e9c1-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-428">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-429">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="1e9c1-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="1e9c1-430">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-430">- TaskPane</span></span><br><span data-ttu-id="1e9c1-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1e9c1-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1e9c1-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1e9c1-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-436">- BindingEvents</span></span><br><span data-ttu-id="1e9c1-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-437">
         - CompressedFile</span></span><br><span data-ttu-id="1e9c1-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1e9c1-438">
         - CustomXmlParts</span></span><br><span data-ttu-id="1e9c1-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-439">
         - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-440">
         - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-440">
         - File</span></span><br><span data-ttu-id="1e9c1-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-441">
         - HtmlCoercion</span></span><br><span data-ttu-id="1e9c1-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-442">
         - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-443">
         - MatrixBindings</span></span><br><span data-ttu-id="1e9c1-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-444">
         - MatrixCoercion</span></span><br><span data-ttu-id="1e9c1-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-445">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1e9c1-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-446">
         - PdfFile</span></span><br><span data-ttu-id="1e9c1-447">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-447">
         - Selection</span></span><br><span data-ttu-id="1e9c1-448">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-448">
         - Settings</span></span><br><span data-ttu-id="1e9c1-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-449">
         - TableBindings</span></span><br><span data-ttu-id="1e9c1-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-450">
         - TableCoercion</span></span><br><span data-ttu-id="1e9c1-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-451">
         - TextBindings</span></span><br><span data-ttu-id="1e9c1-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-452">
         - TextCoercion</span></span><br><span data-ttu-id="1e9c1-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-453">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-454">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="1e9c1-454">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="1e9c1-455">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-455">- TaskPane</span></span><br><span data-ttu-id="1e9c1-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1e9c1-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1e9c1-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1e9c1-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-461">- BindingEvents</span></span><br><span data-ttu-id="1e9c1-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-462">
         - CompressedFile</span></span><br><span data-ttu-id="1e9c1-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1e9c1-463">
         - CustomXmlParts</span></span><br><span data-ttu-id="1e9c1-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-464">
         - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-465">
         - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-465">
         - File</span></span><br><span data-ttu-id="1e9c1-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-466">
         - HtmlCoercion</span></span><br><span data-ttu-id="1e9c1-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-467">
         - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-468">
         - MatrixBindings</span></span><br><span data-ttu-id="1e9c1-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-469">
         - MatrixCoercion</span></span><br><span data-ttu-id="1e9c1-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-470">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1e9c1-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-471">
         - PdfFile</span></span><br><span data-ttu-id="1e9c1-472">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-472">
         - Selection</span></span><br><span data-ttu-id="1e9c1-473">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-473">
         - Settings</span></span><br><span data-ttu-id="1e9c1-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-474">
         - TableBindings</span></span><br><span data-ttu-id="1e9c1-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-475">
         - TableCoercion</span></span><br><span data-ttu-id="1e9c1-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-476">
         - TextBindings</span></span><br><span data-ttu-id="1e9c1-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-477">
         - TextCoercion</span></span><br><span data-ttu-id="1e9c1-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-478">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-479">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="1e9c1-479">Office for iPad</span></span></td>
    <td> <span data-ttu-id="1e9c1-480">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-480">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1e9c1-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1e9c1-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1e9c1-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1e9c1-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1e9c1-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1e9c1-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-485">- BindingEvents</span></span><br><span data-ttu-id="1e9c1-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-486">
         - CompressedFile</span></span><br><span data-ttu-id="1e9c1-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1e9c1-487">
         - CustomXmlParts</span></span><br><span data-ttu-id="1e9c1-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-488">
         - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-489">
         - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-489">
         - File</span></span><br><span data-ttu-id="1e9c1-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-490">
         - HtmlCoercion</span></span><br><span data-ttu-id="1e9c1-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-491">
         - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-492">
         - MatrixBindings</span></span><br><span data-ttu-id="1e9c1-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-493">
         - MatrixCoercion</span></span><br><span data-ttu-id="1e9c1-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-494">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1e9c1-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-495">
         - PdfFile</span></span><br><span data-ttu-id="1e9c1-496">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-496">
         - Selection</span></span><br><span data-ttu-id="1e9c1-497">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-497">
         - Settings</span></span><br><span data-ttu-id="1e9c1-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-498">
         - TableBindings</span></span><br><span data-ttu-id="1e9c1-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-499">
         - TableCoercion</span></span><br><span data-ttu-id="1e9c1-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-500">
         - TextBindings</span></span><br><span data-ttu-id="1e9c1-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-501">
         - TextCoercion</span></span><br><span data-ttu-id="1e9c1-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-502">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-503">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="1e9c1-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="1e9c1-504">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-504">- TaskPane</span></span><br><span data-ttu-id="1e9c1-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1e9c1-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1e9c1-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1e9c1-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1e9c1-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1e9c1-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-510">- BindingEvents</span></span><br><span data-ttu-id="1e9c1-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-511">
         - CompressedFile</span></span><br><span data-ttu-id="1e9c1-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1e9c1-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="1e9c1-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-513">
         - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-514">
         - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-514">
         - File</span></span><br><span data-ttu-id="1e9c1-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="1e9c1-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-516">
         - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-517">
         - MatrixBindings</span></span><br><span data-ttu-id="1e9c1-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-518">
         - MatrixCoercion</span></span><br><span data-ttu-id="1e9c1-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-519">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1e9c1-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-520">
         - PdfFile</span></span><br><span data-ttu-id="1e9c1-521">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-521">
         - Selection</span></span><br><span data-ttu-id="1e9c1-522">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-522">
         - Settings</span></span><br><span data-ttu-id="1e9c1-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-523">
         - TableBindings</span></span><br><span data-ttu-id="1e9c1-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-524">
         - TableCoercion</span></span><br><span data-ttu-id="1e9c1-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-525">
         - TextBindings</span></span><br><span data-ttu-id="1e9c1-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-526">
         - TextCoercion</span></span><br><span data-ttu-id="1e9c1-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-527">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-528">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="1e9c1-528">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="1e9c1-529">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-529">- TaskPane</span></span><br><span data-ttu-id="1e9c1-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1e9c1-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1e9c1-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1e9c1-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1e9c1-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1e9c1-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-535">- BindingEvents</span></span><br><span data-ttu-id="1e9c1-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-536">
         - CompressedFile</span></span><br><span data-ttu-id="1e9c1-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1e9c1-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="1e9c1-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-538">
         - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-539">
         - File</span></span><br><span data-ttu-id="1e9c1-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="1e9c1-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-541">
         - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-542">
         - MatrixBindings</span></span><br><span data-ttu-id="1e9c1-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-543">
         - MatrixCoercion</span></span><br><span data-ttu-id="1e9c1-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-544">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1e9c1-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-545">
         - PdfFile</span></span><br><span data-ttu-id="1e9c1-546">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-546">
         - Selection</span></span><br><span data-ttu-id="1e9c1-547">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-547">
         - Settings</span></span><br><span data-ttu-id="1e9c1-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-548">
         - TableBindings</span></span><br><span data-ttu-id="1e9c1-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-549">
         - TableCoercion</span></span><br><span data-ttu-id="1e9c1-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-550">
         - TextBindings</span></span><br><span data-ttu-id="1e9c1-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-551">
         - TextCoercion</span></span><br><span data-ttu-id="1e9c1-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-552">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="1e9c1-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="1e9c1-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1e9c1-554">平台</span><span class="sxs-lookup"><span data-stu-id="1e9c1-554">Platform</span></span></th>
    <th><span data-ttu-id="1e9c1-555">扩展点</span><span class="sxs-lookup"><span data-stu-id="1e9c1-555">Extension points</span></span></th>
    <th><span data-ttu-id="1e9c1-556">API 要求集</span><span class="sxs-lookup"><span data-stu-id="1e9c1-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="1e9c1-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="1e9c1-558">Office Online</span></span></td>
    <td> <span data-ttu-id="1e9c1-559">- 内容</span><span class="sxs-lookup"><span data-stu-id="1e9c1-559">- Content</span></span><br><span data-ttu-id="1e9c1-560">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-560">
         - TaskPane</span></span><br><span data-ttu-id="1e9c1-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1e9c1-563">- ActiveView</span></span><br><span data-ttu-id="1e9c1-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-564">
         - CompressedFile</span></span><br><span data-ttu-id="1e9c1-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-565">
         - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-566">
         - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-566">
         - File</span></span><br><span data-ttu-id="1e9c1-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-567">
         - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-568">
         - PdfFile</span></span><br><span data-ttu-id="1e9c1-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-569">
         - Selection</span></span><br><span data-ttu-id="1e9c1-570">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-570">
         - Settings</span></span><br><span data-ttu-id="1e9c1-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-571">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-572">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="1e9c1-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="1e9c1-573">- 内容</span><span class="sxs-lookup"><span data-stu-id="1e9c1-573">- Content</span></span><br><span data-ttu-id="1e9c1-574">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-574">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="1e9c1-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1e9c1-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1e9c1-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1e9c1-576">- ActiveView</span></span><br><span data-ttu-id="1e9c1-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-577">
         - CompressedFile</span></span><br><span data-ttu-id="1e9c1-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-578">
         - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-579">
         - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-579">
         - File</span></span><br><span data-ttu-id="1e9c1-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-580">
         - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-581">
         - PdfFile</span></span><br><span data-ttu-id="1e9c1-582">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-582">
         - Selection</span></span><br><span data-ttu-id="1e9c1-583">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-583">
         - Settings</span></span><br><span data-ttu-id="1e9c1-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-584">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-585">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="1e9c1-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="1e9c1-586">- 内容</span><span class="sxs-lookup"><span data-stu-id="1e9c1-586">- Content</span></span><br><span data-ttu-id="1e9c1-587">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-587">
         - TaskPane</span></span><br><span data-ttu-id="1e9c1-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1e9c1-590">- ActiveView</span></span><br><span data-ttu-id="1e9c1-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-591">
         - CompressedFile</span></span><br><span data-ttu-id="1e9c1-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-592">
         - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-593">
         - File</span></span><br><span data-ttu-id="1e9c1-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-594">
         - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-595">
         - PdfFile</span></span><br><span data-ttu-id="1e9c1-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-596">
         - Selection</span></span><br><span data-ttu-id="1e9c1-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-597">
         - Settings</span></span><br><span data-ttu-id="1e9c1-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-598">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-599">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="1e9c1-599">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="1e9c1-600">- 内容</span><span class="sxs-lookup"><span data-stu-id="1e9c1-600">- Content</span></span><br><span data-ttu-id="1e9c1-601">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-601">
         - TaskPane</span></span><br><span data-ttu-id="1e9c1-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1e9c1-604">- ActiveView</span></span><br><span data-ttu-id="1e9c1-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-605">
         - CompressedFile</span></span><br><span data-ttu-id="1e9c1-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-606">
         - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-607">
         - File</span></span><br><span data-ttu-id="1e9c1-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-608">
         - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-609">
         - PdfFile</span></span><br><span data-ttu-id="1e9c1-610">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-610">
         - Selection</span></span><br><span data-ttu-id="1e9c1-611">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-611">
         - Settings</span></span><br><span data-ttu-id="1e9c1-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-612">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-613">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="1e9c1-613">Office for iPad</span></span></td>
    <td> <span data-ttu-id="1e9c1-614">- 内容</span><span class="sxs-lookup"><span data-stu-id="1e9c1-614">- Content</span></span><br><span data-ttu-id="1e9c1-615">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-615">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="1e9c1-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="1e9c1-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1e9c1-617">- ActiveView</span></span><br><span data-ttu-id="1e9c1-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-618">
         - CompressedFile</span></span><br><span data-ttu-id="1e9c1-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-619">
         - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-620">
         - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-620">
         - File</span></span><br><span data-ttu-id="1e9c1-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-621">
         - PdfFile</span></span><br><span data-ttu-id="1e9c1-622">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-622">
         - Selection</span></span><br><span data-ttu-id="1e9c1-623">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-623">
         - Settings</span></span><br><span data-ttu-id="1e9c1-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-624">
         - TextCoercion</span></span><br><span data-ttu-id="1e9c1-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-625">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-626">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="1e9c1-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="1e9c1-627">- 内容</span><span class="sxs-lookup"><span data-stu-id="1e9c1-627">- Content</span></span><br><span data-ttu-id="1e9c1-628">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-628">
         - TaskPane</span></span><br><span data-ttu-id="1e9c1-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1e9c1-631">- ActiveView</span></span><br><span data-ttu-id="1e9c1-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-632">
         - CompressedFile</span></span><br><span data-ttu-id="1e9c1-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-633">
         - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-634">
         - File</span></span><br><span data-ttu-id="1e9c1-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-635">
         - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-636">
         - PdfFile</span></span><br><span data-ttu-id="1e9c1-637">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-637">
         - Selection</span></span><br><span data-ttu-id="1e9c1-638">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-638">
         - Settings</span></span><br><span data-ttu-id="1e9c1-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-639">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-640">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="1e9c1-640">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="1e9c1-641">- 内容</span><span class="sxs-lookup"><span data-stu-id="1e9c1-641">- Content</span></span><br><span data-ttu-id="1e9c1-642">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-642">
         - TaskPane</span></span><br><span data-ttu-id="1e9c1-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1e9c1-645">- ActiveView</span></span><br><span data-ttu-id="1e9c1-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-646">
         - CompressedFile</span></span><br><span data-ttu-id="1e9c1-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-647">
         - DocumentEvents</span></span><br><span data-ttu-id="1e9c1-648">
         - File</span><span class="sxs-lookup"><span data-stu-id="1e9c1-648">
         - File</span></span><br><span data-ttu-id="1e9c1-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-649">
         - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1e9c1-650">
         - PdfFile</span></span><br><span data-ttu-id="1e9c1-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-651">
         - Selection</span></span><br><span data-ttu-id="1e9c1-652">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-652">
         - Settings</span></span><br><span data-ttu-id="1e9c1-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-653">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="1e9c1-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="1e9c1-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1e9c1-655">平台</span><span class="sxs-lookup"><span data-stu-id="1e9c1-655">Platform</span></span></th>
    <th><span data-ttu-id="1e9c1-656">扩展点</span><span class="sxs-lookup"><span data-stu-id="1e9c1-656">Extension points</span></span></th>
    <th><span data-ttu-id="1e9c1-657">API 要求集</span><span class="sxs-lookup"><span data-stu-id="1e9c1-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="1e9c1-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="1e9c1-659">Office Online</span></span></td>
    <td> <span data-ttu-id="1e9c1-660">- 内容</span><span class="sxs-lookup"><span data-stu-id="1e9c1-660">- Content</span></span><br><span data-ttu-id="1e9c1-661">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-661">
         - TaskPane</span></span><br><span data-ttu-id="1e9c1-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="1e9c1-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1e9c1-665">- DocumentEvents</span></span><br><span data-ttu-id="1e9c1-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="1e9c1-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-667">
         - ImageCoercion</span></span><br><span data-ttu-id="1e9c1-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1e9c1-668">
         - Settings</span></span><br><span data-ttu-id="1e9c1-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-669">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="1e9c1-670">项目</span><span class="sxs-lookup"><span data-stu-id="1e9c1-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1e9c1-671">平台</span><span class="sxs-lookup"><span data-stu-id="1e9c1-671">Platform</span></span></th>
    <th><span data-ttu-id="1e9c1-672">扩展点</span><span class="sxs-lookup"><span data-stu-id="1e9c1-672">Extension points</span></span></th>
    <th><span data-ttu-id="1e9c1-673">API 要求集</span><span class="sxs-lookup"><span data-stu-id="1e9c1-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="1e9c1-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-675">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="1e9c1-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="1e9c1-676">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-676">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1e9c1-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-678">- Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-678">- Selection</span></span><br><span data-ttu-id="1e9c1-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-679">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-680">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="1e9c1-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="1e9c1-681">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-681">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1e9c1-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-683">- Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-683">- Selection</span></span><br><span data-ttu-id="1e9c1-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-684">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1e9c1-685">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="1e9c1-685">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="1e9c1-686">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="1e9c1-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1e9c1-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1e9c1-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1e9c1-688">- Selection</span><span class="sxs-lookup"><span data-stu-id="1e9c1-688">- Selection</span></span><br><span data-ttu-id="1e9c1-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1e9c1-689">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="1e9c1-690">另请参阅</span><span class="sxs-lookup"><span data-stu-id="1e9c1-690">See also</span></span>

- [<span data-ttu-id="1e9c1-691">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="1e9c1-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="1e9c1-692">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="1e9c1-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="1e9c1-693">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="1e9c1-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="1e9c1-694">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="1e9c1-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)

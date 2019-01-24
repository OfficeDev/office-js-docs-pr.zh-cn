---
title: Office 外接程序主机和平台可用性
description: Excel、Word、Outlook、PowerPoint、OneNote 和项目支持的要求集。
ms.date: 11/07/2018
localization_priority: Priority
ms.openlocfilehash: 9f8b94483d22f24dcb0a6a2ad99df6167533133f
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388337"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="b4edd-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="b4edd-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="b4edd-104">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。</span><span class="sxs-lookup"><span data-stu-id="b4edd-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="b4edd-105">下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="b4edd-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="b4edd-p102">通过 MSI 安装的 Office 2016 的生成号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="b4edd-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="b4edd-108">Excel</span><span class="sxs-lookup"><span data-stu-id="b4edd-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b4edd-109">平台</span><span class="sxs-lookup"><span data-stu-id="b4edd-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b4edd-110">扩展点</span><span class="sxs-lookup"><span data-stu-id="b4edd-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b4edd-111">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b4edd-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b4edd-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b4edd-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="b4edd-113">Office Online</span></span></td>
    <td> <span data-ttu-id="b4edd-114">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-114">- TaskPane</span></span><br><span data-ttu-id="b4edd-115">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b4edd-115">
        - Content</span></span><br><span data-ttu-id="b4edd-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="b4edd-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b4edd-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b4edd-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b4edd-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b4edd-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b4edd-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b4edd-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b4edd-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b4edd-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b4edd-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b4edd-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-126">
        - BindingEvents</span></span><br><span data-ttu-id="b4edd-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-127">
        - CompressedFile</span></span><br><span data-ttu-id="b4edd-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-128">
        - DocumentEvents</span></span><br><span data-ttu-id="b4edd-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-129">
        - File</span></span><br><span data-ttu-id="b4edd-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-130">
        - MatrixBindings</span></span><br><span data-ttu-id="b4edd-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="b4edd-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-132">
        - Selection</span></span><br><span data-ttu-id="b4edd-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-133">
        - Settings</span></span><br><span data-ttu-id="b4edd-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-134">
        - TableBindings</span></span><br><span data-ttu-id="b4edd-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-135">
        - TableCoercion</span></span><br><span data-ttu-id="b4edd-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-136">
        - TextBindings</span></span><br><span data-ttu-id="b4edd-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-138">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="b4edd-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="b4edd-139">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-139">
        - TaskPane</span></span><br><span data-ttu-id="b4edd-140">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b4edd-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="b4edd-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b4edd-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-142">
        - BindingEvents</span></span><br><span data-ttu-id="b4edd-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-143">
        - CompressedFile</span></span><br><span data-ttu-id="b4edd-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-144">
        - DocumentEvents</span></span><br><span data-ttu-id="b4edd-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-145">
        - File</span></span><br><span data-ttu-id="b4edd-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-146">
        - ImageCoercion</span></span><br><span data-ttu-id="b4edd-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-147">
        - MatrixBindings</span></span><br><span data-ttu-id="b4edd-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="b4edd-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-149">
        - Selection</span></span><br><span data-ttu-id="b4edd-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-150">
        - Settings</span></span><br><span data-ttu-id="b4edd-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-151">
        - TableBindings</span></span><br><span data-ttu-id="b4edd-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-152">
        - TableCoercion</span></span><br><span data-ttu-id="b4edd-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-153">
        - TextBindings</span></span><br><span data-ttu-id="b4edd-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-155">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="b4edd-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="b4edd-156">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-156">- TaskPane</span></span><br><span data-ttu-id="b4edd-157">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b4edd-157">
        - Content</span></span><br><span data-ttu-id="b4edd-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b4edd-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b4edd-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b4edd-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b4edd-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b4edd-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b4edd-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b4edd-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b4edd-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b4edd-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b4edd-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-168">- BindingEvents</span></span><br><span data-ttu-id="b4edd-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-169">
        - CompressedFile</span></span><br><span data-ttu-id="b4edd-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-170">
        - DocumentEvents</span></span><br><span data-ttu-id="b4edd-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-171">
        - File</span></span><br><span data-ttu-id="b4edd-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-172">
        - ImageCoercion</span></span><br><span data-ttu-id="b4edd-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-173">
        - MatrixBindings</span></span><br><span data-ttu-id="b4edd-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="b4edd-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-175">
        - Selection</span></span><br><span data-ttu-id="b4edd-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-176">
        - Settings</span></span><br><span data-ttu-id="b4edd-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-177">
        - TableBindings</span></span><br><span data-ttu-id="b4edd-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-178">
        - TableCoercion</span></span><br><span data-ttu-id="b4edd-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-179">
        - TextBindings</span></span><br><span data-ttu-id="b4edd-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-181">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="b4edd-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="b4edd-182">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-182">- TaskPane</span></span><br><span data-ttu-id="b4edd-183">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b4edd-183">
        - Content</span></span><br><span data-ttu-id="b4edd-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b4edd-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b4edd-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b4edd-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b4edd-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b4edd-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b4edd-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b4edd-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b4edd-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b4edd-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b4edd-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-194">- BindingEvents</span></span><br><span data-ttu-id="b4edd-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-195">
        - CompressedFile</span></span><br><span data-ttu-id="b4edd-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-196">
        - DocumentEvents</span></span><br><span data-ttu-id="b4edd-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-197">
        - File</span></span><br><span data-ttu-id="b4edd-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-198">
        - ImageCoercion</span></span><br><span data-ttu-id="b4edd-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-199">
        - MatrixBindings</span></span><br><span data-ttu-id="b4edd-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="b4edd-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-201">
        - Selection</span></span><br><span data-ttu-id="b4edd-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-202">
        - Settings</span></span><br><span data-ttu-id="b4edd-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-203">
        - TableBindings</span></span><br><span data-ttu-id="b4edd-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-204">
        - TableCoercion</span></span><br><span data-ttu-id="b4edd-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-205">
        - TextBindings</span></span><br><span data-ttu-id="b4edd-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-207">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="b4edd-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="b4edd-208">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-208">- TaskPane</span></span><br><span data-ttu-id="b4edd-209">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b4edd-209">
        - Content</span></span></td>
    <td><span data-ttu-id="b4edd-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b4edd-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b4edd-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b4edd-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b4edd-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b4edd-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b4edd-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b4edd-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b4edd-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b4edd-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-219">- BindingEvents</span></span><br><span data-ttu-id="b4edd-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-220">
        - CompressedFile</span></span><br><span data-ttu-id="b4edd-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-221">
        - DocumentEvents</span></span><br><span data-ttu-id="b4edd-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-222">
        - File</span></span><br><span data-ttu-id="b4edd-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-223">
        - ImageCoercion</span></span><br><span data-ttu-id="b4edd-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-224">
        - MatrixBindings</span></span><br><span data-ttu-id="b4edd-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="b4edd-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-226">
        - Selection</span></span><br><span data-ttu-id="b4edd-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-227">
        - Settings</span></span><br><span data-ttu-id="b4edd-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-228">
        - TableBindings</span></span><br><span data-ttu-id="b4edd-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-229">
        - TableCoercion</span></span><br><span data-ttu-id="b4edd-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-230">
        - TextBindings</span></span><br><span data-ttu-id="b4edd-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-232">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="b4edd-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="b4edd-233">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-233">- TaskPane</span></span><br><span data-ttu-id="b4edd-234">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b4edd-234">
        - Content</span></span><br><span data-ttu-id="b4edd-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b4edd-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b4edd-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b4edd-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b4edd-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b4edd-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b4edd-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b4edd-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b4edd-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b4edd-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b4edd-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-245">- BindingEvents</span></span><br><span data-ttu-id="b4edd-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-246">
        - CompressedFile</span></span><br><span data-ttu-id="b4edd-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-247">
        - DocumentEvents</span></span><br><span data-ttu-id="b4edd-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-248">
        - File</span></span><br><span data-ttu-id="b4edd-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-249">
        - ImageCoercion</span></span><br><span data-ttu-id="b4edd-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-250">
        - MatrixBindings</span></span><br><span data-ttu-id="b4edd-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="b4edd-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-252">
        - PdfFile</span></span><br><span data-ttu-id="b4edd-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-253">
        - Selection</span></span><br><span data-ttu-id="b4edd-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-254">
        - Settings</span></span><br><span data-ttu-id="b4edd-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-255">
        - TableBindings</span></span><br><span data-ttu-id="b4edd-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-256">
        - TableCoercion</span></span><br><span data-ttu-id="b4edd-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-257">
        - TextBindings</span></span><br><span data-ttu-id="b4edd-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-259">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="b4edd-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="b4edd-260">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-260">- TaskPane</span></span><br><span data-ttu-id="b4edd-261">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b4edd-261">
        - Content</span></span><br><span data-ttu-id="b4edd-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b4edd-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b4edd-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b4edd-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b4edd-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b4edd-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b4edd-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b4edd-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b4edd-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b4edd-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b4edd-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-272">- BindingEvents</span></span><br><span data-ttu-id="b4edd-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-273">
        - CompressedFile</span></span><br><span data-ttu-id="b4edd-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-274">
        - DocumentEvents</span></span><br><span data-ttu-id="b4edd-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-275">
        - File</span></span><br><span data-ttu-id="b4edd-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-276">
        - ImageCoercion</span></span><br><span data-ttu-id="b4edd-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-277">
        - MatrixBindings</span></span><br><span data-ttu-id="b4edd-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="b4edd-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-279">
        - PdfFile</span></span><br><span data-ttu-id="b4edd-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-280">
        - Selection</span></span><br><span data-ttu-id="b4edd-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-281">
        - Settings</span></span><br><span data-ttu-id="b4edd-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-282">
        - TableBindings</span></span><br><span data-ttu-id="b4edd-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-283">
        - TableCoercion</span></span><br><span data-ttu-id="b4edd-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-284">
        - TextBindings</span></span><br><span data-ttu-id="b4edd-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="b4edd-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="b4edd-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b4edd-287">平台</span><span class="sxs-lookup"><span data-stu-id="b4edd-287">Platform</span></span></th>
    <th><span data-ttu-id="b4edd-288">扩展点</span><span class="sxs-lookup"><span data-stu-id="b4edd-288">Extension points</span></span></th>
    <th><span data-ttu-id="b4edd-289">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b4edd-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="b4edd-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b4edd-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="b4edd-291">Office Online</span></span></td>
    <td> <span data-ttu-id="b4edd-292">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b4edd-292">- Mail Read</span></span><br><span data-ttu-id="b4edd-293">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="b4edd-293">
      - Mail Compose</span></span><br><span data-ttu-id="b4edd-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b4edd-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b4edd-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b4edd-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b4edd-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b4edd-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b4edd-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b4edd-302">不可用</span><span class="sxs-lookup"><span data-stu-id="b4edd-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-303">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="b4edd-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b4edd-304">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b4edd-304">- Mail Read</span></span><br><span data-ttu-id="b4edd-305">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="b4edd-305">
      - Mail Compose</span></span><br><span data-ttu-id="b4edd-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b4edd-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b4edd-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b4edd-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="b4edd-311">不可用</span><span class="sxs-lookup"><span data-stu-id="b4edd-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-312">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="b4edd-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b4edd-313">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b4edd-313">- Mail Read</span></span><br><span data-ttu-id="b4edd-314">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="b4edd-314">
      - Mail Compose</span></span><br><span data-ttu-id="b4edd-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b4edd-316">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="b4edd-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b4edd-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b4edd-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b4edd-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b4edd-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b4edd-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b4edd-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b4edd-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b4edd-324">不可用</span><span class="sxs-lookup"><span data-stu-id="b4edd-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-325">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="b4edd-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="b4edd-326">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b4edd-326">- Mail Read</span></span><br><span data-ttu-id="b4edd-327">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="b4edd-327">
      - Mail Compose</span></span><br><span data-ttu-id="b4edd-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b4edd-329">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="b4edd-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b4edd-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b4edd-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b4edd-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b4edd-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b4edd-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b4edd-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b4edd-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b4edd-337">不可用</span><span class="sxs-lookup"><span data-stu-id="b4edd-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-338">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="b4edd-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="b4edd-339">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b4edd-339">- Mail Read</span></span><br><span data-ttu-id="b4edd-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b4edd-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b4edd-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b4edd-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b4edd-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b4edd-346">不可用</span><span class="sxs-lookup"><span data-stu-id="b4edd-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-347">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="b4edd-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="b4edd-348">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b4edd-348">- Mail Read</span></span><br><span data-ttu-id="b4edd-349">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="b4edd-349">
      - Mail Compose</span></span><br><span data-ttu-id="b4edd-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b4edd-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b4edd-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b4edd-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b4edd-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b4edd-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b4edd-357">不可用</span><span class="sxs-lookup"><span data-stu-id="b4edd-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-358">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="b4edd-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="b4edd-359">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b4edd-359">- Mail Read</span></span><br><span data-ttu-id="b4edd-360">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="b4edd-360">
      - Mail Compose</span></span><br><span data-ttu-id="b4edd-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b4edd-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b4edd-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b4edd-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b4edd-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b4edd-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b4edd-368">不可用</span><span class="sxs-lookup"><span data-stu-id="b4edd-368">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-369">Office for Android</span><span class="sxs-lookup"><span data-stu-id="b4edd-369">Office for Android</span></span></td>
    <td> <span data-ttu-id="b4edd-370">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b4edd-370">- Mail Read</span></span><br><span data-ttu-id="b4edd-371">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-371">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b4edd-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b4edd-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b4edd-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b4edd-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b4edd-377">不可用</span><span class="sxs-lookup"><span data-stu-id="b4edd-377">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="b4edd-378">Word</span><span class="sxs-lookup"><span data-stu-id="b4edd-378">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b4edd-379">平台</span><span class="sxs-lookup"><span data-stu-id="b4edd-379">Platform</span></span></th>
    <th><span data-ttu-id="b4edd-380">扩展点</span><span class="sxs-lookup"><span data-stu-id="b4edd-380">Extension points</span></span></th>
    <th><span data-ttu-id="b4edd-381">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b4edd-381">API requirement sets</span></span></th>
    <th><span data-ttu-id="b4edd-382"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b4edd-382"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-383">Office Online</span><span class="sxs-lookup"><span data-stu-id="b4edd-383">Office Online</span></span></td>
    <td> <span data-ttu-id="b4edd-384">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-384">- TaskPane</span></span><br><span data-ttu-id="b4edd-385">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-385">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-386">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-386">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b4edd-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b4edd-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b4edd-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b4edd-390">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-390">- BindingEvents</span></span><br><span data-ttu-id="b4edd-391">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b4edd-391">
         - CustomXmlParts</span></span><br><span data-ttu-id="b4edd-392">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-392">
         - DocumentEvents</span></span><br><span data-ttu-id="b4edd-393">
         - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-393">
         - File</span></span><br><span data-ttu-id="b4edd-394">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-394">
         - HtmlCoercion</span></span><br><span data-ttu-id="b4edd-395">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-395">
         - ImageCoercion</span></span><br><span data-ttu-id="b4edd-396">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-396">
         - MatrixBindings</span></span><br><span data-ttu-id="b4edd-397">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-397">
         - MatrixCoercion</span></span><br><span data-ttu-id="b4edd-398">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-398">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b4edd-399">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-399">
         - PdfFile</span></span><br><span data-ttu-id="b4edd-400">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-400">
         - Selection</span></span><br><span data-ttu-id="b4edd-401">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-401">
         - Settings</span></span><br><span data-ttu-id="b4edd-402">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-402">
         - TableBindings</span></span><br><span data-ttu-id="b4edd-403">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-403">
         - TableCoercion</span></span><br><span data-ttu-id="b4edd-404">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-404">
         - TextBindings</span></span><br><span data-ttu-id="b4edd-405">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-405">
         - TextCoercion</span></span><br><span data-ttu-id="b4edd-406">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-406">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-407">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="b4edd-407">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b4edd-408">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-408">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b4edd-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b4edd-410">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-410">- BindingEvents</span></span><br><span data-ttu-id="b4edd-411">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-411">
         - CompressedFile</span></span><br><span data-ttu-id="b4edd-412">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b4edd-412">
         - CustomXmlParts</span></span><br><span data-ttu-id="b4edd-413">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-413">
         - DocumentEvents</span></span><br><span data-ttu-id="b4edd-414">
         - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-414">
         - File</span></span><br><span data-ttu-id="b4edd-415">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-415">
         - HtmlCoercion</span></span><br><span data-ttu-id="b4edd-416">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-416">
         - ImageCoercion</span></span><br><span data-ttu-id="b4edd-417">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-417">
         - MatrixBindings</span></span><br><span data-ttu-id="b4edd-418">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-418">
         - MatrixCoercion</span></span><br><span data-ttu-id="b4edd-419">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-419">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b4edd-420">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-420">
         - PdfFile</span></span><br><span data-ttu-id="b4edd-421">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-421">
         - Selection</span></span><br><span data-ttu-id="b4edd-422">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-422">
         - Settings</span></span><br><span data-ttu-id="b4edd-423">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-423">
         - TableBindings</span></span><br><span data-ttu-id="b4edd-424">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-424">
         - TableCoercion</span></span><br><span data-ttu-id="b4edd-425">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-425">
         - TextBindings</span></span><br><span data-ttu-id="b4edd-426">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-426">
         - TextCoercion</span></span><br><span data-ttu-id="b4edd-427">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-427">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-428">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="b4edd-428">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b4edd-429">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-429">- TaskPane</span></span><br><span data-ttu-id="b4edd-430">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-430">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-431">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-431">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b4edd-432">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-432">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b4edd-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b4edd-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b4edd-435">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-435">- BindingEvents</span></span><br><span data-ttu-id="b4edd-436">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-436">
         - CompressedFile</span></span><br><span data-ttu-id="b4edd-437">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b4edd-437">
         - CustomXmlParts</span></span><br><span data-ttu-id="b4edd-438">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-438">
         - DocumentEvents</span></span><br><span data-ttu-id="b4edd-439">
         - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-439">
         - File</span></span><br><span data-ttu-id="b4edd-440">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-440">
         - HtmlCoercion</span></span><br><span data-ttu-id="b4edd-441">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-441">
         - ImageCoercion</span></span><br><span data-ttu-id="b4edd-442">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-442">
         - MatrixBindings</span></span><br><span data-ttu-id="b4edd-443">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-443">
         - MatrixCoercion</span></span><br><span data-ttu-id="b4edd-444">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-444">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b4edd-445">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-445">
         - PdfFile</span></span><br><span data-ttu-id="b4edd-446">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-446">
         - Selection</span></span><br><span data-ttu-id="b4edd-447">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-447">
         - Settings</span></span><br><span data-ttu-id="b4edd-448">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-448">
         - TableBindings</span></span><br><span data-ttu-id="b4edd-449">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-449">
         - TableCoercion</span></span><br><span data-ttu-id="b4edd-450">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-450">
         - TextBindings</span></span><br><span data-ttu-id="b4edd-451">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-451">
         - TextCoercion</span></span><br><span data-ttu-id="b4edd-452">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-452">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-453">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="b4edd-453">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="b4edd-454">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-454">- TaskPane</span></span><br><span data-ttu-id="b4edd-455">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-455">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-456">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-456">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b4edd-457">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-457">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b4edd-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b4edd-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b4edd-460">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-460">- BindingEvents</span></span><br><span data-ttu-id="b4edd-461">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-461">
         - CompressedFile</span></span><br><span data-ttu-id="b4edd-462">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b4edd-462">
         - CustomXmlParts</span></span><br><span data-ttu-id="b4edd-463">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-463">
         - DocumentEvents</span></span><br><span data-ttu-id="b4edd-464">
         - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-464">
         - File</span></span><br><span data-ttu-id="b4edd-465">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-465">
         - HtmlCoercion</span></span><br><span data-ttu-id="b4edd-466">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-466">
         - ImageCoercion</span></span><br><span data-ttu-id="b4edd-467">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-467">
         - MatrixBindings</span></span><br><span data-ttu-id="b4edd-468">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-468">
         - MatrixCoercion</span></span><br><span data-ttu-id="b4edd-469">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-469">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b4edd-470">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-470">
         - PdfFile</span></span><br><span data-ttu-id="b4edd-471">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-471">
         - Selection</span></span><br><span data-ttu-id="b4edd-472">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-472">
         - Settings</span></span><br><span data-ttu-id="b4edd-473">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-473">
         - TableBindings</span></span><br><span data-ttu-id="b4edd-474">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-474">
         - TableCoercion</span></span><br><span data-ttu-id="b4edd-475">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-475">
         - TextBindings</span></span><br><span data-ttu-id="b4edd-476">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-476">
         - TextCoercion</span></span><br><span data-ttu-id="b4edd-477">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-477">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-478">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="b4edd-478">Office for iPad</span></span></td>
    <td> <span data-ttu-id="b4edd-479">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-479">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b4edd-480">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-480">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b4edd-481">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-481">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b4edd-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b4edd-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b4edd-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b4edd-484">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-484">- BindingEvents</span></span><br><span data-ttu-id="b4edd-485">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-485">
         - CompressedFile</span></span><br><span data-ttu-id="b4edd-486">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b4edd-486">
         - CustomXmlParts</span></span><br><span data-ttu-id="b4edd-487">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-487">
         - DocumentEvents</span></span><br><span data-ttu-id="b4edd-488">
         - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-488">
         - File</span></span><br><span data-ttu-id="b4edd-489">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-489">
         - HtmlCoercion</span></span><br><span data-ttu-id="b4edd-490">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-490">
         - ImageCoercion</span></span><br><span data-ttu-id="b4edd-491">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-491">
         - MatrixBindings</span></span><br><span data-ttu-id="b4edd-492">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-492">
         - MatrixCoercion</span></span><br><span data-ttu-id="b4edd-493">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-493">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b4edd-494">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-494">
         - PdfFile</span></span><br><span data-ttu-id="b4edd-495">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-495">
         - Selection</span></span><br><span data-ttu-id="b4edd-496">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-496">
         - Settings</span></span><br><span data-ttu-id="b4edd-497">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-497">
         - TableBindings</span></span><br><span data-ttu-id="b4edd-498">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-498">
         - TableCoercion</span></span><br><span data-ttu-id="b4edd-499">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-499">
         - TextBindings</span></span><br><span data-ttu-id="b4edd-500">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-500">
         - TextCoercion</span></span><br><span data-ttu-id="b4edd-501">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-501">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-502">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="b4edd-502">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="b4edd-503">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-503">- TaskPane</span></span><br><span data-ttu-id="b4edd-504">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-504">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-505">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-505">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b4edd-506">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-506">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b4edd-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b4edd-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b4edd-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b4edd-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-509">- BindingEvents</span></span><br><span data-ttu-id="b4edd-510">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-510">
         - CompressedFile</span></span><br><span data-ttu-id="b4edd-511">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b4edd-511">
         - CustomXmlParts</span></span><br><span data-ttu-id="b4edd-512">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-512">
         - DocumentEvents</span></span><br><span data-ttu-id="b4edd-513">
         - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-513">
         - File</span></span><br><span data-ttu-id="b4edd-514">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-514">
         - HtmlCoercion</span></span><br><span data-ttu-id="b4edd-515">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-515">
         - ImageCoercion</span></span><br><span data-ttu-id="b4edd-516">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-516">
         - MatrixBindings</span></span><br><span data-ttu-id="b4edd-517">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-517">
         - MatrixCoercion</span></span><br><span data-ttu-id="b4edd-518">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-518">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b4edd-519">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-519">
         - PdfFile</span></span><br><span data-ttu-id="b4edd-520">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-520">
         - Selection</span></span><br><span data-ttu-id="b4edd-521">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-521">
         - Settings</span></span><br><span data-ttu-id="b4edd-522">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-522">
         - TableBindings</span></span><br><span data-ttu-id="b4edd-523">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-523">
         - TableCoercion</span></span><br><span data-ttu-id="b4edd-524">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-524">
         - TextBindings</span></span><br><span data-ttu-id="b4edd-525">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-525">
         - TextCoercion</span></span><br><span data-ttu-id="b4edd-526">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-526">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-527">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="b4edd-527">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="b4edd-528">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-528">- TaskPane</span></span><br><span data-ttu-id="b4edd-529">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-529">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-530">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-530">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b4edd-531">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-531">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b4edd-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b4edd-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b4edd-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b4edd-534">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-534">- BindingEvents</span></span><br><span data-ttu-id="b4edd-535">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-535">
         - CompressedFile</span></span><br><span data-ttu-id="b4edd-536">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b4edd-536">
         - CustomXmlParts</span></span><br><span data-ttu-id="b4edd-537">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-537">
         - DocumentEvents</span></span><br><span data-ttu-id="b4edd-538">
         - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-538">
         - File</span></span><br><span data-ttu-id="b4edd-539">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-539">
         - HtmlCoercion</span></span><br><span data-ttu-id="b4edd-540">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-540">
         - ImageCoercion</span></span><br><span data-ttu-id="b4edd-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-541">
         - MatrixBindings</span></span><br><span data-ttu-id="b4edd-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="b4edd-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b4edd-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-544">
         - PdfFile</span></span><br><span data-ttu-id="b4edd-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-545">
         - Selection</span></span><br><span data-ttu-id="b4edd-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-546">
         - Settings</span></span><br><span data-ttu-id="b4edd-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-547">
         - TableBindings</span></span><br><span data-ttu-id="b4edd-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-548">
         - TableCoercion</span></span><br><span data-ttu-id="b4edd-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b4edd-549">
         - TextBindings</span></span><br><span data-ttu-id="b4edd-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-550">
         - TextCoercion</span></span><br><span data-ttu-id="b4edd-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-551">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="b4edd-552">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b4edd-552">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b4edd-553">平台</span><span class="sxs-lookup"><span data-stu-id="b4edd-553">Platform</span></span></th>
    <th><span data-ttu-id="b4edd-554">扩展点</span><span class="sxs-lookup"><span data-stu-id="b4edd-554">Extension points</span></span></th>
    <th><span data-ttu-id="b4edd-555">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b4edd-555">API requirement sets</span></span></th>
    <th><span data-ttu-id="b4edd-556"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b4edd-556"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-557">Office Online</span><span class="sxs-lookup"><span data-stu-id="b4edd-557">Office Online</span></span></td>
    <td> <span data-ttu-id="b4edd-558">- 内容</span><span class="sxs-lookup"><span data-stu-id="b4edd-558">- Content</span></span><br><span data-ttu-id="b4edd-559">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-559">
         - TaskPane</span></span><br><span data-ttu-id="b4edd-560">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-560">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b4edd-562">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b4edd-562">- ActiveView</span></span><br><span data-ttu-id="b4edd-563">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-563">
         - CompressedFile</span></span><br><span data-ttu-id="b4edd-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-564">
         - DocumentEvents</span></span><br><span data-ttu-id="b4edd-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-565">
         - File</span></span><br><span data-ttu-id="b4edd-566">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-566">
         - ImageCoercion</span></span><br><span data-ttu-id="b4edd-567">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-567">
         - PdfFile</span></span><br><span data-ttu-id="b4edd-568">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-568">
         - Selection</span></span><br><span data-ttu-id="b4edd-569">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-569">
         - Settings</span></span><br><span data-ttu-id="b4edd-570">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-570">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-571">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="b4edd-571">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b4edd-572">- 内容</span><span class="sxs-lookup"><span data-stu-id="b4edd-572">- Content</span></span><br><span data-ttu-id="b4edd-573">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-573">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="b4edd-574">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b4edd-574">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b4edd-575">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b4edd-575">- ActiveView</span></span><br><span data-ttu-id="b4edd-576">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-576">
         - CompressedFile</span></span><br><span data-ttu-id="b4edd-577">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-577">
         - DocumentEvents</span></span><br><span data-ttu-id="b4edd-578">
         - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-578">
         - File</span></span><br><span data-ttu-id="b4edd-579">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-579">
         - ImageCoercion</span></span><br><span data-ttu-id="b4edd-580">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-580">
         - PdfFile</span></span><br><span data-ttu-id="b4edd-581">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-581">
         - Selection</span></span><br><span data-ttu-id="b4edd-582">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-582">
         - Settings</span></span><br><span data-ttu-id="b4edd-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-583">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-584">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="b4edd-584">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b4edd-585">- 内容</span><span class="sxs-lookup"><span data-stu-id="b4edd-585">- Content</span></span><br><span data-ttu-id="b4edd-586">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-586">
         - TaskPane</span></span><br><span data-ttu-id="b4edd-587">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-587">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-588">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-588">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b4edd-589">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b4edd-589">- ActiveView</span></span><br><span data-ttu-id="b4edd-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-590">
         - CompressedFile</span></span><br><span data-ttu-id="b4edd-591">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-591">
         - DocumentEvents</span></span><br><span data-ttu-id="b4edd-592">
         - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-592">
         - File</span></span><br><span data-ttu-id="b4edd-593">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-593">
         - ImageCoercion</span></span><br><span data-ttu-id="b4edd-594">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-594">
         - PdfFile</span></span><br><span data-ttu-id="b4edd-595">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-595">
         - Selection</span></span><br><span data-ttu-id="b4edd-596">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-596">
         - Settings</span></span><br><span data-ttu-id="b4edd-597">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-597">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-598">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="b4edd-598">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="b4edd-599">- 内容</span><span class="sxs-lookup"><span data-stu-id="b4edd-599">- Content</span></span><br><span data-ttu-id="b4edd-600">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-600">
         - TaskPane</span></span><br><span data-ttu-id="b4edd-601">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-601">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-602">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-602">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b4edd-603">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b4edd-603">- ActiveView</span></span><br><span data-ttu-id="b4edd-604">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-604">
         - CompressedFile</span></span><br><span data-ttu-id="b4edd-605">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-605">
         - DocumentEvents</span></span><br><span data-ttu-id="b4edd-606">
         - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-606">
         - File</span></span><br><span data-ttu-id="b4edd-607">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-607">
         - ImageCoercion</span></span><br><span data-ttu-id="b4edd-608">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-608">
         - PdfFile</span></span><br><span data-ttu-id="b4edd-609">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-609">
         - Selection</span></span><br><span data-ttu-id="b4edd-610">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-610">
         - Settings</span></span><br><span data-ttu-id="b4edd-611">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-611">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-612">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="b4edd-612">Office for iPad</span></span></td>
    <td> <span data-ttu-id="b4edd-613">- 内容</span><span class="sxs-lookup"><span data-stu-id="b4edd-613">- Content</span></span><br><span data-ttu-id="b4edd-614">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-614">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b4edd-615">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-615">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="b4edd-616">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b4edd-616">- ActiveView</span></span><br><span data-ttu-id="b4edd-617">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-617">
         - CompressedFile</span></span><br><span data-ttu-id="b4edd-618">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-618">
         - DocumentEvents</span></span><br><span data-ttu-id="b4edd-619">
         - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-619">
         - File</span></span><br><span data-ttu-id="b4edd-620">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-620">
         - PdfFile</span></span><br><span data-ttu-id="b4edd-621">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-621">
         - Selection</span></span><br><span data-ttu-id="b4edd-622">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-622">
         - Settings</span></span><br><span data-ttu-id="b4edd-623">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-623">
         - TextCoercion</span></span><br><span data-ttu-id="b4edd-624">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-624">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-625">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="b4edd-625">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="b4edd-626">- 内容</span><span class="sxs-lookup"><span data-stu-id="b4edd-626">- Content</span></span><br><span data-ttu-id="b4edd-627">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-627">
         - TaskPane</span></span><br><span data-ttu-id="b4edd-628">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-628">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-629">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-629">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b4edd-630">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b4edd-630">- ActiveView</span></span><br><span data-ttu-id="b4edd-631">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-631">
         - CompressedFile</span></span><br><span data-ttu-id="b4edd-632">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-632">
         - DocumentEvents</span></span><br><span data-ttu-id="b4edd-633">
         - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-633">
         - File</span></span><br><span data-ttu-id="b4edd-634">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-634">
         - ImageCoercion</span></span><br><span data-ttu-id="b4edd-635">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-635">
         - PdfFile</span></span><br><span data-ttu-id="b4edd-636">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-636">
         - Selection</span></span><br><span data-ttu-id="b4edd-637">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-637">
         - Settings</span></span><br><span data-ttu-id="b4edd-638">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-638">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-639">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="b4edd-639">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="b4edd-640">- 内容</span><span class="sxs-lookup"><span data-stu-id="b4edd-640">- Content</span></span><br><span data-ttu-id="b4edd-641">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-641">
         - TaskPane</span></span><br><span data-ttu-id="b4edd-642">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-642">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-643">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-643">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b4edd-644">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b4edd-644">- ActiveView</span></span><br><span data-ttu-id="b4edd-645">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-645">
         - CompressedFile</span></span><br><span data-ttu-id="b4edd-646">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-646">
         - DocumentEvents</span></span><br><span data-ttu-id="b4edd-647">
         - File</span><span class="sxs-lookup"><span data-stu-id="b4edd-647">
         - File</span></span><br><span data-ttu-id="b4edd-648">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-648">
         - ImageCoercion</span></span><br><span data-ttu-id="b4edd-649">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b4edd-649">
         - PdfFile</span></span><br><span data-ttu-id="b4edd-650">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-650">
         - Selection</span></span><br><span data-ttu-id="b4edd-651">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-651">
         - Settings</span></span><br><span data-ttu-id="b4edd-652">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-652">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="b4edd-653">OneNote</span><span class="sxs-lookup"><span data-stu-id="b4edd-653">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b4edd-654">平台</span><span class="sxs-lookup"><span data-stu-id="b4edd-654">Platform</span></span></th>
    <th><span data-ttu-id="b4edd-655">扩展点</span><span class="sxs-lookup"><span data-stu-id="b4edd-655">Extension points</span></span></th>
    <th><span data-ttu-id="b4edd-656">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b4edd-656">API requirement sets</span></span></th>
    <th><span data-ttu-id="b4edd-657"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b4edd-657"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-658">Office Online</span><span class="sxs-lookup"><span data-stu-id="b4edd-658">Office Online</span></span></td>
    <td> <span data-ttu-id="b4edd-659">- 内容</span><span class="sxs-lookup"><span data-stu-id="b4edd-659">- Content</span></span><br><span data-ttu-id="b4edd-660">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-660">
         - TaskPane</span></span><br><span data-ttu-id="b4edd-661">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-661">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b4edd-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="b4edd-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b4edd-664">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b4edd-664">- DocumentEvents</span></span><br><span data-ttu-id="b4edd-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="b4edd-666">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-666">
         - ImageCoercion</span></span><br><span data-ttu-id="b4edd-667">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b4edd-667">
         - Settings</span></span><br><span data-ttu-id="b4edd-668">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-668">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="b4edd-669">项目</span><span class="sxs-lookup"><span data-stu-id="b4edd-669">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b4edd-670">平台</span><span class="sxs-lookup"><span data-stu-id="b4edd-670">Platform</span></span></th>
    <th><span data-ttu-id="b4edd-671">扩展点</span><span class="sxs-lookup"><span data-stu-id="b4edd-671">Extension points</span></span></th>
    <th><span data-ttu-id="b4edd-672">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b4edd-672">API requirement sets</span></span></th>
    <th><span data-ttu-id="b4edd-673"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b4edd-673"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-674">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="b4edd-674">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b4edd-675">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-675">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b4edd-676">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-676">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b4edd-677">- Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-677">- Selection</span></span><br><span data-ttu-id="b4edd-678">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-678">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-679">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="b4edd-679">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b4edd-680">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-680">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b4edd-681">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-681">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b4edd-682">- Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-682">- Selection</span></span><br><span data-ttu-id="b4edd-683">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-683">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b4edd-684">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="b4edd-684">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="b4edd-685">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b4edd-685">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b4edd-686">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b4edd-686">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b4edd-687">- Selection</span><span class="sxs-lookup"><span data-stu-id="b4edd-687">- Selection</span></span><br><span data-ttu-id="b4edd-688">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b4edd-688">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="b4edd-689">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b4edd-689">See also</span></span>

- [<span data-ttu-id="b4edd-690">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="b4edd-690">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="b4edd-691">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="b4edd-691">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="b4edd-692">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="b4edd-692">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="b4edd-693">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="b4edd-693">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)

---
title: Office 外接程序主机和平台可用性
description: Excel、Word、Outlook、PowerPoint、OneNote 和项目支持的要求集。
ms.date: 03/15/2019
localization_priority: Priority
ms.openlocfilehash: 4348881c35e4c79975d34406e4668b2693405134
ms.sourcegitcommit: c4d6ecdc41ea67291b6d155c3b246e31ec2e38b7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/16/2019
ms.locfileid: "30654961"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="b9b24-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="b9b24-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="b9b24-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="b9b24-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="b9b24-p102">通过 MSI 安装的 Office 2016 的生成号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="b9b24-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>
>
> <span data-ttu-id="b9b24-108">Office 2019 的一次性购买的内部版本号是 16.0.10827.20150。</span><span class="sxs-lookup"><span data-stu-id="b9b24-108">The build number for a one-time purchase of Office 2019 is 16.0.10827.20150.</span></span>

## <a name="excel"></a><span data-ttu-id="b9b24-109">Excel</span><span class="sxs-lookup"><span data-stu-id="b9b24-109">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b9b24-110">平台</span><span class="sxs-lookup"><span data-stu-id="b9b24-110">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b9b24-111">扩展点</span><span class="sxs-lookup"><span data-stu-id="b9b24-111">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b9b24-112">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b9b24-112">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b9b24-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b9b24-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="b9b24-114">Office Online</span></span></td>
    <td> <span data-ttu-id="b9b24-115">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-115">- TaskPane</span></span><br><span data-ttu-id="b9b24-116">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-116">
        - Content</span></span><br><span data-ttu-id="b9b24-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="b9b24-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b9b24-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b9b24-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b9b24-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b9b24-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b9b24-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b9b24-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b9b24-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b9b24-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b9b24-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b9b24-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-127">
        - BindingEvents</span></span><br><span data-ttu-id="b9b24-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-128">
        - CompressedFile</span></span><br><span data-ttu-id="b9b24-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-129">
        - DocumentEvents</span></span><br><span data-ttu-id="b9b24-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-130">
        - File</span></span><br><span data-ttu-id="b9b24-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-131">
        - MatrixBindings</span></span><br><span data-ttu-id="b9b24-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-132">
        - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-133">
        - Selection</span></span><br><span data-ttu-id="b9b24-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-134">
        - Settings</span></span><br><span data-ttu-id="b9b24-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-135">
        - TableBindings</span></span><br><span data-ttu-id="b9b24-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-136">
        - TableCoercion</span></span><br><span data-ttu-id="b9b24-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-137">
        - TextBindings</span></span><br><span data-ttu-id="b9b24-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-138">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-139">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-139">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-140">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-140">- TaskPane</span></span><br><span data-ttu-id="b9b24-141">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-141">
        - Content</span></span><br><span data-ttu-id="b9b24-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="b9b24-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b9b24-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b9b24-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b9b24-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b9b24-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b9b24-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b9b24-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b9b24-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b9b24-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b9b24-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b9b24-152">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-152">
        - BindingEvents</span></span><br><span data-ttu-id="b9b24-153">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-153">
        - CompressedFile</span></span><br><span data-ttu-id="b9b24-154">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-154">
        - DocumentEvents</span></span><br><span data-ttu-id="b9b24-155">
        - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-155">
        - File</span></span><br><span data-ttu-id="b9b24-156">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-156">
        - MatrixBindings</span></span><br><span data-ttu-id="b9b24-157">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-157">
        - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-158">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-158">
        - Selection</span></span><br><span data-ttu-id="b9b24-159">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-159">
        - Settings</span></span><br><span data-ttu-id="b9b24-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-160">
        - TableBindings</span></span><br><span data-ttu-id="b9b24-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-161">
        - TableCoercion</span></span><br><span data-ttu-id="b9b24-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-162">
        - TextBindings</span></span><br><span data-ttu-id="b9b24-163">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-163">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-164">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-164">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="b9b24-165">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-165">- TaskPane</span></span><br><span data-ttu-id="b9b24-166">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-166">
        - Content</span></span><br><span data-ttu-id="b9b24-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b9b24-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b9b24-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b9b24-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b9b24-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b9b24-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b9b24-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b9b24-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b9b24-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b9b24-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b9b24-177">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-177">- BindingEvents</span></span><br><span data-ttu-id="b9b24-178">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-178">
        - CompressedFile</span></span><br><span data-ttu-id="b9b24-179">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-179">
        - DocumentEvents</span></span><br><span data-ttu-id="b9b24-180">
        - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-180">
        - File</span></span><br><span data-ttu-id="b9b24-181">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-181">
        - ImageCoercion</span></span><br><span data-ttu-id="b9b24-182">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-182">
        - MatrixBindings</span></span><br><span data-ttu-id="b9b24-183">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-183">
        - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-184">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-184">
        - Selection</span></span><br><span data-ttu-id="b9b24-185">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-185">
        - Settings</span></span><br><span data-ttu-id="b9b24-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-186">
        - TableBindings</span></span><br><span data-ttu-id="b9b24-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-187">
        - TableCoercion</span></span><br><span data-ttu-id="b9b24-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-188">
        - TextBindings</span></span><br><span data-ttu-id="b9b24-189">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-189">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-190">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-190">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="b9b24-191">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-191">- TaskPane</span></span><br><span data-ttu-id="b9b24-192">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-192">
        - Content</span></span></td>
    <td><span data-ttu-id="b9b24-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b9b24-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b9b24-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="b9b24-195">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-195">- BindingEvents</span></span><br><span data-ttu-id="b9b24-196">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-196">
        - CompressedFile</span></span><br><span data-ttu-id="b9b24-197">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-197">
        - DocumentEvents</span></span><br><span data-ttu-id="b9b24-198">
        - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-198">
        - File</span></span><br><span data-ttu-id="b9b24-199">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-199">
        - ImageCoercion</span></span><br><span data-ttu-id="b9b24-200">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-200">
        - MatrixBindings</span></span><br><span data-ttu-id="b9b24-201">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-201">
        - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-202">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-202">
        - Selection</span></span><br><span data-ttu-id="b9b24-203">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-203">
        - Settings</span></span><br><span data-ttu-id="b9b24-204">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-204">
        - TableBindings</span></span><br><span data-ttu-id="b9b24-205">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-205">
        - TableCoercion</span></span><br><span data-ttu-id="b9b24-206">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-206">
        - TextBindings</span></span><br><span data-ttu-id="b9b24-207">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-207">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-208">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-208">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="b9b24-209">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-209">
        - TaskPane</span></span><br><span data-ttu-id="b9b24-210">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-210">
        - Content</span></span></td>
    <td>  <span data-ttu-id="b9b24-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b9b24-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="b9b24-212">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-212">
        - BindingEvents</span></span><br><span data-ttu-id="b9b24-213">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-213">
        - CompressedFile</span></span><br><span data-ttu-id="b9b24-214">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-214">
        - DocumentEvents</span></span><br><span data-ttu-id="b9b24-215">
        - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-215">
        - File</span></span><br><span data-ttu-id="b9b24-216">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-216">
        - ImageCoercion</span></span><br><span data-ttu-id="b9b24-217">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-217">
        - MatrixBindings</span></span><br><span data-ttu-id="b9b24-218">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-218">
        - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-219">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-219">
        - Selection</span></span><br><span data-ttu-id="b9b24-220">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-220">
        - Settings</span></span><br><span data-ttu-id="b9b24-221">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-221">
        - TableBindings</span></span><br><span data-ttu-id="b9b24-222">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-222">
        - TableCoercion</span></span><br><span data-ttu-id="b9b24-223">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-223">
        - TextBindings</span></span><br><span data-ttu-id="b9b24-224">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-224">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-225">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="b9b24-225">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="b9b24-226">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-226">- TaskPane</span></span><br><span data-ttu-id="b9b24-227">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-227">
        - Content</span></span></td>
    <td><span data-ttu-id="b9b24-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b9b24-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b9b24-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b9b24-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b9b24-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b9b24-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b9b24-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b9b24-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b9b24-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b9b24-237">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-237">- BindingEvents</span></span><br><span data-ttu-id="b9b24-238">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-238">
        - CompressedFile</span></span><br><span data-ttu-id="b9b24-239">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-239">
        - DocumentEvents</span></span><br><span data-ttu-id="b9b24-240">
        - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-240">
        - File</span></span><br><span data-ttu-id="b9b24-241">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-241">
        - ImageCoercion</span></span><br><span data-ttu-id="b9b24-242">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-242">
        - MatrixBindings</span></span><br><span data-ttu-id="b9b24-243">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-243">
        - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-244">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-244">
        - Selection</span></span><br><span data-ttu-id="b9b24-245">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-245">
        - Settings</span></span><br><span data-ttu-id="b9b24-246">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-246">
        - TableBindings</span></span><br><span data-ttu-id="b9b24-247">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-247">
        - TableCoercion</span></span><br><span data-ttu-id="b9b24-248">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-248">
        - TextBindings</span></span><br><span data-ttu-id="b9b24-249">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-249">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-250">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="b9b24-250">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="b9b24-251">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-251">- TaskPane</span></span><br><span data-ttu-id="b9b24-252">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-252">
        - Content</span></span><br><span data-ttu-id="b9b24-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b9b24-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b9b24-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b9b24-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b9b24-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b9b24-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b9b24-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b9b24-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b9b24-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b9b24-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b9b24-263">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-263">- BindingEvents</span></span><br><span data-ttu-id="b9b24-264">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-264">
        - CompressedFile</span></span><br><span data-ttu-id="b9b24-265">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-265">
        - DocumentEvents</span></span><br><span data-ttu-id="b9b24-266">
        - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-266">
        - File</span></span><br><span data-ttu-id="b9b24-267">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-267">
        - ImageCoercion</span></span><br><span data-ttu-id="b9b24-268">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-268">
        - MatrixBindings</span></span><br><span data-ttu-id="b9b24-269">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-269">
        - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-270">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-270">
        - PdfFile</span></span><br><span data-ttu-id="b9b24-271">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-271">
        - Selection</span></span><br><span data-ttu-id="b9b24-272">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-272">
        - Settings</span></span><br><span data-ttu-id="b9b24-273">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-273">
        - TableBindings</span></span><br><span data-ttu-id="b9b24-274">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-274">
        - TableCoercion</span></span><br><span data-ttu-id="b9b24-275">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-275">
        - TextBindings</span></span><br><span data-ttu-id="b9b24-276">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-276">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-277">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="b9b24-277">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="b9b24-278">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-278">- TaskPane</span></span><br><span data-ttu-id="b9b24-279">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-279">
        - Content</span></span><br><span data-ttu-id="b9b24-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b9b24-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b9b24-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b9b24-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b9b24-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b9b24-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b9b24-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b9b24-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b9b24-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b9b24-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b9b24-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-290">- BindingEvents</span></span><br><span data-ttu-id="b9b24-291">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-291">
        - CompressedFile</span></span><br><span data-ttu-id="b9b24-292">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-292">
        - DocumentEvents</span></span><br><span data-ttu-id="b9b24-293">
        - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-293">
        - File</span></span><br><span data-ttu-id="b9b24-294">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-294">
        - ImageCoercion</span></span><br><span data-ttu-id="b9b24-295">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-295">
        - MatrixBindings</span></span><br><span data-ttu-id="b9b24-296">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-296">
        - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-297">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-297">
        - PdfFile</span></span><br><span data-ttu-id="b9b24-298">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-298">
        - Selection</span></span><br><span data-ttu-id="b9b24-299">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-299">
        - Settings</span></span><br><span data-ttu-id="b9b24-300">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-300">
        - TableBindings</span></span><br><span data-ttu-id="b9b24-301">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-301">
        - TableCoercion</span></span><br><span data-ttu-id="b9b24-302">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-302">
        - TextBindings</span></span><br><span data-ttu-id="b9b24-303">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-303">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-304">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="b9b24-304">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="b9b24-305">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-305">- TaskPane</span></span><br><span data-ttu-id="b9b24-306">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-306">
        - Content</span></span></td>
    <td><span data-ttu-id="b9b24-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b9b24-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b9b24-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="b9b24-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-309">- BindingEvents</span></span><br><span data-ttu-id="b9b24-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-310">
        - CompressedFile</span></span><br><span data-ttu-id="b9b24-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-311">
        - DocumentEvents</span></span><br><span data-ttu-id="b9b24-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-312">
        - File</span></span><br><span data-ttu-id="b9b24-313">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-313">
        - ImageCoercion</span></span><br><span data-ttu-id="b9b24-314">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-314">
        - MatrixBindings</span></span><br><span data-ttu-id="b9b24-315">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-315">
        - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-316">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-316">
        - PdfFile</span></span><br><span data-ttu-id="b9b24-317">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-317">
        - Selection</span></span><br><span data-ttu-id="b9b24-318">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-318">
        - Settings</span></span><br><span data-ttu-id="b9b24-319">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-319">
        - TableBindings</span></span><br><span data-ttu-id="b9b24-320">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-320">
        - TableCoercion</span></span><br><span data-ttu-id="b9b24-321">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-321">
        - TextBindings</span></span><br><span data-ttu-id="b9b24-322">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-322">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b9b24-323">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="b9b24-323">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="b9b24-324">Outlook</span><span class="sxs-lookup"><span data-stu-id="b9b24-324">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b9b24-325">平台</span><span class="sxs-lookup"><span data-stu-id="b9b24-325">Platform</span></span></th>
    <th><span data-ttu-id="b9b24-326">扩展点</span><span class="sxs-lookup"><span data-stu-id="b9b24-326">Extension points</span></span></th>
    <th><span data-ttu-id="b9b24-327">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b9b24-327">API requirement sets</span></span></th>
    <th><span data-ttu-id="b9b24-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b9b24-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-329">Office Online</span><span class="sxs-lookup"><span data-stu-id="b9b24-329">Office Online</span></span></td>
    <td> <span data-ttu-id="b9b24-330">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b9b24-330">- Mail Read</span></span><br><span data-ttu-id="b9b24-331">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="b9b24-331">
      - Mail Compose</span></span><br><span data-ttu-id="b9b24-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b9b24-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b9b24-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b9b24-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b9b24-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b9b24-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b9b24-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b9b24-340">不可用</span><span class="sxs-lookup"><span data-stu-id="b9b24-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-341">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-341">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-342">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b9b24-342">- Mail Read</span></span><br><span data-ttu-id="b9b24-343">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="b9b24-343">
      - Mail Compose</span></span><br><span data-ttu-id="b9b24-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b9b24-345">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="b9b24-345">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b9b24-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b9b24-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b9b24-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b9b24-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b9b24-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b9b24-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b9b24-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b9b24-353">不可用</span><span class="sxs-lookup"><span data-stu-id="b9b24-353">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-354">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-354">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-355">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b9b24-355">- Mail Read</span></span><br><span data-ttu-id="b9b24-356">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="b9b24-356">
      - Mail Compose</span></span><br><span data-ttu-id="b9b24-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b9b24-358">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="b9b24-358">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b9b24-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b9b24-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b9b24-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b9b24-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b9b24-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b9b24-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b9b24-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b9b24-366">不可用</span><span class="sxs-lookup"><span data-stu-id="b9b24-366">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-367">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-367">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-368">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b9b24-368">- Mail Read</span></span><br><span data-ttu-id="b9b24-369">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="b9b24-369">
      - Mail Compose</span></span><br><span data-ttu-id="b9b24-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b9b24-371">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="b9b24-371">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b9b24-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b9b24-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b9b24-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b9b24-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="b9b24-376">不可用</span><span class="sxs-lookup"><span data-stu-id="b9b24-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-377">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-377">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-378">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b9b24-378">- Mail Read</span></span><br><span data-ttu-id="b9b24-379">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="b9b24-379">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="b9b24-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b9b24-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b9b24-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b9b24-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="b9b24-384">不可用</span><span class="sxs-lookup"><span data-stu-id="b9b24-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-385">Office 365 for iOS</span><span class="sxs-lookup"><span data-stu-id="b9b24-385">See the Office 365 SDK for iOS.</span></span></td>
    <td> <span data-ttu-id="b9b24-386">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b9b24-386">- Mail Read</span></span><br><span data-ttu-id="b9b24-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b9b24-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b9b24-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b9b24-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b9b24-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b9b24-393">不可用</span><span class="sxs-lookup"><span data-stu-id="b9b24-393">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-394">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="b9b24-394">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="b9b24-395">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b9b24-395">- Mail Read</span></span><br><span data-ttu-id="b9b24-396">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="b9b24-396">
      - Mail Compose</span></span><br><span data-ttu-id="b9b24-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b9b24-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b9b24-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b9b24-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b9b24-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b9b24-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b9b24-404">不可用</span><span class="sxs-lookup"><span data-stu-id="b9b24-404">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-405">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="b9b24-405">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="b9b24-406">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b9b24-406">- Mail Read</span></span><br><span data-ttu-id="b9b24-407">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="b9b24-407">
      - Mail Compose</span></span><br><span data-ttu-id="b9b24-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b9b24-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b9b24-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b9b24-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b9b24-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b9b24-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b9b24-415">不可用</span><span class="sxs-lookup"><span data-stu-id="b9b24-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-416">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="b9b24-416">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="b9b24-417">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b9b24-417">- Mail Read</span></span><br><span data-ttu-id="b9b24-418">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="b9b24-418">
      - Mail Compose</span></span><br><span data-ttu-id="b9b24-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b9b24-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b9b24-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b9b24-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b9b24-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b9b24-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b9b24-426">不可用</span><span class="sxs-lookup"><span data-stu-id="b9b24-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-427">Office 365 for Android</span><span class="sxs-lookup"><span data-stu-id="b9b24-427">See the Office 365 SDK for Android.</span></span></td>
    <td> <span data-ttu-id="b9b24-428">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="b9b24-428">- Mail Read</span></span><br><span data-ttu-id="b9b24-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b9b24-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b9b24-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b9b24-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b9b24-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b9b24-435">不可用</span><span class="sxs-lookup"><span data-stu-id="b9b24-435">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="b9b24-436">Word</span><span class="sxs-lookup"><span data-stu-id="b9b24-436">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b9b24-437">平台</span><span class="sxs-lookup"><span data-stu-id="b9b24-437">Platform</span></span></th>
    <th><span data-ttu-id="b9b24-438">扩展点</span><span class="sxs-lookup"><span data-stu-id="b9b24-438">Extension points</span></span></th>
    <th><span data-ttu-id="b9b24-439">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b9b24-439">API requirement sets</span></span></th>
    <th><span data-ttu-id="b9b24-440"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b9b24-440"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-441">Office Online</span><span class="sxs-lookup"><span data-stu-id="b9b24-441">Office Online</span></span></td>
    <td> <span data-ttu-id="b9b24-442">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-442">- TaskPane</span></span><br><span data-ttu-id="b9b24-443">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-443">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-444">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-444">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b9b24-445">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-445">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b9b24-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b9b24-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b9b24-448">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-448">- BindingEvents</span></span><br><span data-ttu-id="b9b24-449">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b9b24-449">
         - CustomXmlParts</span></span><br><span data-ttu-id="b9b24-450">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-450">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-451">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-451">
         - File</span></span><br><span data-ttu-id="b9b24-452">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-452">
         - HtmlCoercion</span></span><br><span data-ttu-id="b9b24-453">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-453">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-454">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-454">
         - MatrixBindings</span></span><br><span data-ttu-id="b9b24-455">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-455">
         - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-456">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-456">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b9b24-457">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-457">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-458">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-458">
         - Selection</span></span><br><span data-ttu-id="b9b24-459">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-459">
         - Settings</span></span><br><span data-ttu-id="b9b24-460">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-460">
         - TableBindings</span></span><br><span data-ttu-id="b9b24-461">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-461">
         - TableCoercion</span></span><br><span data-ttu-id="b9b24-462">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-462">
         - TextBindings</span></span><br><span data-ttu-id="b9b24-463">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-463">
         - TextCoercion</span></span><br><span data-ttu-id="b9b24-464">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-464">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-465">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-465">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-466">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-466">- TaskPane</span></span><br><span data-ttu-id="b9b24-467">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-467">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-468">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-468">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b9b24-469">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-469">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b9b24-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b9b24-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b9b24-472">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-472">- BindingEvents</span></span><br><span data-ttu-id="b9b24-473">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-473">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-474">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b9b24-474">
         - CustomXmlParts</span></span><br><span data-ttu-id="b9b24-475">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-475">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-476">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-476">
         - File</span></span><br><span data-ttu-id="b9b24-477">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-477">
         - HtmlCoercion</span></span><br><span data-ttu-id="b9b24-478">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-478">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-479">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-479">
         - MatrixBindings</span></span><br><span data-ttu-id="b9b24-480">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-480">
         - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-481">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-481">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b9b24-482">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-482">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-483">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-483">
         - Selection</span></span><br><span data-ttu-id="b9b24-484">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-484">
         - Settings</span></span><br><span data-ttu-id="b9b24-485">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-485">
         - TableBindings</span></span><br><span data-ttu-id="b9b24-486">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-486">
         - TableCoercion</span></span><br><span data-ttu-id="b9b24-487">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-487">
         - TextBindings</span></span><br><span data-ttu-id="b9b24-488">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-488">
         - TextCoercion</span></span><br><span data-ttu-id="b9b24-489">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-489">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-490">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-490">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-491">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-491">- TaskPane</span></span><br><span data-ttu-id="b9b24-492">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-492">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-493">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-493">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b9b24-494">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-494">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b9b24-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b9b24-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b9b24-497">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-497">- BindingEvents</span></span><br><span data-ttu-id="b9b24-498">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-498">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-499">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b9b24-499">
         - CustomXmlParts</span></span><br><span data-ttu-id="b9b24-500">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-500">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-501">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-501">
         - File</span></span><br><span data-ttu-id="b9b24-502">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-502">
         - HtmlCoercion</span></span><br><span data-ttu-id="b9b24-503">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-503">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-504">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-504">
         - MatrixBindings</span></span><br><span data-ttu-id="b9b24-505">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-505">
         - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-506">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-506">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b9b24-507">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-507">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-508">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-508">
         - Selection</span></span><br><span data-ttu-id="b9b24-509">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-509">
         - Settings</span></span><br><span data-ttu-id="b9b24-510">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-510">
         - TableBindings</span></span><br><span data-ttu-id="b9b24-511">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-511">
         - TableCoercion</span></span><br><span data-ttu-id="b9b24-512">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-512">
         - TextBindings</span></span><br><span data-ttu-id="b9b24-513">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-513">
         - TextCoercion</span></span><br><span data-ttu-id="b9b24-514">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-514">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-515">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-515">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-516">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-516">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b9b24-517">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-517">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b9b24-518">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b9b24-518">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="b9b24-519">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-519">- BindingEvents</span></span><br><span data-ttu-id="b9b24-520">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-520">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-521">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b9b24-521">
         - CustomXmlParts</span></span><br><span data-ttu-id="b9b24-522">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-522">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-523">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-523">
         - File</span></span><br><span data-ttu-id="b9b24-524">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-524">
         - HtmlCoercion</span></span><br><span data-ttu-id="b9b24-525">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-525">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-526">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-526">
         - MatrixBindings</span></span><br><span data-ttu-id="b9b24-527">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-527">
         - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-528">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-528">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b9b24-529">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-529">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-530">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-530">
         - Selection</span></span><br><span data-ttu-id="b9b24-531">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-531">
         - Settings</span></span><br><span data-ttu-id="b9b24-532">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-532">
         - TableBindings</span></span><br><span data-ttu-id="b9b24-533">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-533">
         - TableCoercion</span></span><br><span data-ttu-id="b9b24-534">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-534">
         - TextBindings</span></span><br><span data-ttu-id="b9b24-535">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-535">
         - TextCoercion</span></span><br><span data-ttu-id="b9b24-536">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-536">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-537">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-537">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-538">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-538">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b9b24-539">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b9b24-539">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="b9b24-540">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-540">- BindingEvents</span></span><br><span data-ttu-id="b9b24-541">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-541">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-542">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b9b24-542">
         - CustomXmlParts</span></span><br><span data-ttu-id="b9b24-543">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-543">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-544">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-544">
         - File</span></span><br><span data-ttu-id="b9b24-545">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-545">
         - HtmlCoercion</span></span><br><span data-ttu-id="b9b24-546">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-546">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-547">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-547">
         - MatrixBindings</span></span><br><span data-ttu-id="b9b24-548">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-548">
         - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-549">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-549">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b9b24-550">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-550">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-551">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-551">
         - Selection</span></span><br><span data-ttu-id="b9b24-552">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-552">
         - Settings</span></span><br><span data-ttu-id="b9b24-553">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-553">
         - TableBindings</span></span><br><span data-ttu-id="b9b24-554">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-554">
         - TableCoercion</span></span><br><span data-ttu-id="b9b24-555">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-555">
         - TextBindings</span></span><br><span data-ttu-id="b9b24-556">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-556">
         - TextCoercion</span></span><br><span data-ttu-id="b9b24-557">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-557">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-558">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="b9b24-558">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="b9b24-559">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-559">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b9b24-560">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-560">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b9b24-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b9b24-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b9b24-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b9b24-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b9b24-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-564">- BindingEvents</span></span><br><span data-ttu-id="b9b24-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-565">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b9b24-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="b9b24-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-567">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-568">
         - File</span></span><br><span data-ttu-id="b9b24-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="b9b24-570">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-570">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-571">
         - MatrixBindings</span></span><br><span data-ttu-id="b9b24-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b9b24-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-574">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-575">
         - Selection</span></span><br><span data-ttu-id="b9b24-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-576">
         - Settings</span></span><br><span data-ttu-id="b9b24-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-577">
         - TableBindings</span></span><br><span data-ttu-id="b9b24-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-578">
         - TableCoercion</span></span><br><span data-ttu-id="b9b24-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-579">
         - TextBindings</span></span><br><span data-ttu-id="b9b24-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-580">
         - TextCoercion</span></span><br><span data-ttu-id="b9b24-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-582">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="b9b24-582">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="b9b24-583">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-583">- TaskPane</span></span><br><span data-ttu-id="b9b24-584">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-584">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-585">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-585">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b9b24-586">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-586">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b9b24-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b9b24-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b9b24-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b9b24-589">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-589">- BindingEvents</span></span><br><span data-ttu-id="b9b24-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-590">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-591">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b9b24-591">
         - CustomXmlParts</span></span><br><span data-ttu-id="b9b24-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-592">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-593">
         - File</span></span><br><span data-ttu-id="b9b24-594">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-594">
         - HtmlCoercion</span></span><br><span data-ttu-id="b9b24-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-595">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-596">
         - MatrixBindings</span></span><br><span data-ttu-id="b9b24-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b9b24-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-599">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-600">
         - Selection</span></span><br><span data-ttu-id="b9b24-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-601">
         - Settings</span></span><br><span data-ttu-id="b9b24-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-602">
         - TableBindings</span></span><br><span data-ttu-id="b9b24-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-603">
         - TableCoercion</span></span><br><span data-ttu-id="b9b24-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-604">
         - TextBindings</span></span><br><span data-ttu-id="b9b24-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-605">
         - TextCoercion</span></span><br><span data-ttu-id="b9b24-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-607">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="b9b24-607">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="b9b24-608">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-608">- TaskPane</span></span><br><span data-ttu-id="b9b24-609">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-609">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-610">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-610">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b9b24-611">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-611">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b9b24-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b9b24-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b9b24-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b9b24-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-614">- BindingEvents</span></span><br><span data-ttu-id="b9b24-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-615">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b9b24-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="b9b24-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-617">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-618">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-618">
         - File</span></span><br><span data-ttu-id="b9b24-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="b9b24-620">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-620">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-621">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-621">
         - MatrixBindings</span></span><br><span data-ttu-id="b9b24-622">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-622">
         - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-623">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-623">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b9b24-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-624">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-625">
         - Selection</span></span><br><span data-ttu-id="b9b24-626">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-626">
         - Settings</span></span><br><span data-ttu-id="b9b24-627">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-627">
         - TableBindings</span></span><br><span data-ttu-id="b9b24-628">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-628">
         - TableCoercion</span></span><br><span data-ttu-id="b9b24-629">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-629">
         - TextBindings</span></span><br><span data-ttu-id="b9b24-630">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-630">
         - TextCoercion</span></span><br><span data-ttu-id="b9b24-631">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-631">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-632">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="b9b24-632">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="b9b24-633">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b9b24-634">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-634">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b9b24-635">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b9b24-635">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="b9b24-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-636">- BindingEvents</span></span><br><span data-ttu-id="b9b24-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-637">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b9b24-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="b9b24-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-639">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-640">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-640">
         - File</span></span><br><span data-ttu-id="b9b24-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="b9b24-642">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-642">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-643">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-643">
         - MatrixBindings</span></span><br><span data-ttu-id="b9b24-644">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-644">
         - MatrixCoercion</span></span><br><span data-ttu-id="b9b24-645">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-645">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b9b24-646">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-646">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-647">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-647">
         - Selection</span></span><br><span data-ttu-id="b9b24-648">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-648">
         - Settings</span></span><br><span data-ttu-id="b9b24-649">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-649">
         - TableBindings</span></span><br><span data-ttu-id="b9b24-650">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-650">
         - TableCoercion</span></span><br><span data-ttu-id="b9b24-651">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b9b24-651">
         - TextBindings</span></span><br><span data-ttu-id="b9b24-652">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-652">
         - TextCoercion</span></span><br><span data-ttu-id="b9b24-653">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-653">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="b9b24-654">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="b9b24-654">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="b9b24-655">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b9b24-655">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b9b24-656">平台</span><span class="sxs-lookup"><span data-stu-id="b9b24-656">Platform</span></span></th>
    <th><span data-ttu-id="b9b24-657">扩展点</span><span class="sxs-lookup"><span data-stu-id="b9b24-657">Extension points</span></span></th>
    <th><span data-ttu-id="b9b24-658">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b9b24-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="b9b24-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b9b24-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="b9b24-660">Office Online</span></span></td>
    <td> <span data-ttu-id="b9b24-661">- 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-661">- Content</span></span><br><span data-ttu-id="b9b24-662">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-662">
         - TaskPane</span></span><br><span data-ttu-id="b9b24-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b9b24-665">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b9b24-665">- ActiveView</span></span><br><span data-ttu-id="b9b24-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-666">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-667">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-667">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-668">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-668">
         - File</span></span><br><span data-ttu-id="b9b24-669">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-669">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-670">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-671">
         - Selection</span></span><br><span data-ttu-id="b9b24-672">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-672">
         - Settings</span></span><br><span data-ttu-id="b9b24-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-673">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-674">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-674">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-675">- 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-675">- Content</span></span><br><span data-ttu-id="b9b24-676">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-676">
         - TaskPane</span></span><br><span data-ttu-id="b9b24-677">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-677">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-678">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-678">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b9b24-679">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b9b24-679">- ActiveView</span></span><br><span data-ttu-id="b9b24-680">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-680">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-681">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-681">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-682">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-682">
         - File</span></span><br><span data-ttu-id="b9b24-683">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-683">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-684">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-684">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-685">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-685">
         - Selection</span></span><br><span data-ttu-id="b9b24-686">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-686">
         - Settings</span></span><br><span data-ttu-id="b9b24-687">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-687">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-688">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-688">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-689">- 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-689">- Content</span></span><br><span data-ttu-id="b9b24-690">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-690">
         - TaskPane</span></span><br><span data-ttu-id="b9b24-691">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-691">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-692">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-692">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b9b24-693">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b9b24-693">- ActiveView</span></span><br><span data-ttu-id="b9b24-694">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-694">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-695">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-695">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-696">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-696">
         - File</span></span><br><span data-ttu-id="b9b24-697">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-697">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-698">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-698">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-699">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-699">
         - Selection</span></span><br><span data-ttu-id="b9b24-700">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-700">
         - Settings</span></span><br><span data-ttu-id="b9b24-701">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-701">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-702">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-702">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-703">- 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-703">- Content</span></span><br><span data-ttu-id="b9b24-704">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-704">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b9b24-705">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b9b24-705">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="b9b24-706">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b9b24-706">- ActiveView</span></span><br><span data-ttu-id="b9b24-707">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-707">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-708">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-708">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-709">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-709">
         - File</span></span><br><span data-ttu-id="b9b24-710">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-710">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-711">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-711">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-712">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-712">
         - Selection</span></span><br><span data-ttu-id="b9b24-713">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-713">
         - Settings</span></span><br><span data-ttu-id="b9b24-714">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-714">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-715">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-715">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-716">- 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-716">- Content</span></span><br><span data-ttu-id="b9b24-717">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-717">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="b9b24-718">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b9b24-718">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="b9b24-719">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b9b24-719">- ActiveView</span></span><br><span data-ttu-id="b9b24-720">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-720">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-721">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-721">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-722">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-722">
         - File</span></span><br><span data-ttu-id="b9b24-723">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-723">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-724">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-724">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-725">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-725">
         - Selection</span></span><br><span data-ttu-id="b9b24-726">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-726">
         - Settings</span></span><br><span data-ttu-id="b9b24-727">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-727">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-728">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="b9b24-728">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="b9b24-729">- 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-729">- Content</span></span><br><span data-ttu-id="b9b24-730">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-730">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b9b24-731">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-731">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="b9b24-732">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b9b24-732">- ActiveView</span></span><br><span data-ttu-id="b9b24-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-733">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-734">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-734">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-735">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-735">
         - File</span></span><br><span data-ttu-id="b9b24-736">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-736">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-737">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-737">
         - Selection</span></span><br><span data-ttu-id="b9b24-738">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-738">
         - Settings</span></span><br><span data-ttu-id="b9b24-739">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-739">
         - TextCoercion</span></span><br><span data-ttu-id="b9b24-740">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-740">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-741">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="b9b24-741">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="b9b24-742">- 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-742">- Content</span></span><br><span data-ttu-id="b9b24-743">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-743">
         - TaskPane</span></span><br><span data-ttu-id="b9b24-744">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-744">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-745">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-745">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b9b24-746">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b9b24-746">- ActiveView</span></span><br><span data-ttu-id="b9b24-747">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-747">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-748">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-748">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-749">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-749">
         - File</span></span><br><span data-ttu-id="b9b24-750">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-750">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-751">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-751">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-752">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-752">
         - Selection</span></span><br><span data-ttu-id="b9b24-753">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-753">
         - Settings</span></span><br><span data-ttu-id="b9b24-754">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-754">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-755">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="b9b24-755">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="b9b24-756">- 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-756">- Content</span></span><br><span data-ttu-id="b9b24-757">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-757">
         - TaskPane</span></span><br><span data-ttu-id="b9b24-758">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-758">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-759">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-759">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b9b24-760">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b9b24-760">- ActiveView</span></span><br><span data-ttu-id="b9b24-761">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-761">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-762">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-762">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-763">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-763">
         - File</span></span><br><span data-ttu-id="b9b24-764">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-764">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-765">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-765">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-766">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-766">
         - Selection</span></span><br><span data-ttu-id="b9b24-767">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-767">
         - Settings</span></span><br><span data-ttu-id="b9b24-768">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-768">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-769">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="b9b24-769">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="b9b24-770">- 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-770">- Content</span></span><br><span data-ttu-id="b9b24-771">
         - 任务窗格/td></span><span class="sxs-lookup"><span data-stu-id="b9b24-771">
         - TaskPane/td></span></span> <td> <span data-ttu-id="b9b24-772">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b9b24-772">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="b9b24-773">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b9b24-773">- ActiveView</span></span><br><span data-ttu-id="b9b24-774">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-774">
         - CompressedFile</span></span><br><span data-ttu-id="b9b24-775">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-775">
         - DocumentEvents</span></span><br><span data-ttu-id="b9b24-776">
         - File</span><span class="sxs-lookup"><span data-stu-id="b9b24-776">
         - File</span></span><br><span data-ttu-id="b9b24-777">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-777">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-778">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b9b24-778">
         - PdfFile</span></span><br><span data-ttu-id="b9b24-779">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-779">
         - Selection</span></span><br><span data-ttu-id="b9b24-780">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-780">
         - Settings</span></span><br><span data-ttu-id="b9b24-781">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-781">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b9b24-782">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="b9b24-782">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="b9b24-783">OneNote</span><span class="sxs-lookup"><span data-stu-id="b9b24-783">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b9b24-784">平台</span><span class="sxs-lookup"><span data-stu-id="b9b24-784">Platform</span></span></th>
    <th><span data-ttu-id="b9b24-785">扩展点</span><span class="sxs-lookup"><span data-stu-id="b9b24-785">Extension points</span></span></th>
    <th><span data-ttu-id="b9b24-786">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b9b24-786">API requirement sets</span></span></th>
    <th><span data-ttu-id="b9b24-787"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b9b24-787"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-788">Office Online</span><span class="sxs-lookup"><span data-stu-id="b9b24-788">Office Online</span></span></td>
    <td> <span data-ttu-id="b9b24-789">- 内容</span><span class="sxs-lookup"><span data-stu-id="b9b24-789">- Content</span></span><br><span data-ttu-id="b9b24-790">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-790">
         - TaskPane</span></span><br><span data-ttu-id="b9b24-791">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-791">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b9b24-792">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-792">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="b9b24-793">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-793">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b9b24-794">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b9b24-794">- DocumentEvents</span></span><br><span data-ttu-id="b9b24-795">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-795">
         - HtmlCoercion</span></span><br><span data-ttu-id="b9b24-796">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-796">
         - ImageCoercion</span></span><br><span data-ttu-id="b9b24-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b9b24-797">
         - Settings</span></span><br><span data-ttu-id="b9b24-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-798">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="b9b24-799">项目</span><span class="sxs-lookup"><span data-stu-id="b9b24-799">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b9b24-800">平台</span><span class="sxs-lookup"><span data-stu-id="b9b24-800">Platform</span></span></th>
    <th><span data-ttu-id="b9b24-801">扩展点</span><span class="sxs-lookup"><span data-stu-id="b9b24-801">Extension points</span></span></th>
    <th><span data-ttu-id="b9b24-802">API 要求集</span><span class="sxs-lookup"><span data-stu-id="b9b24-802">API requirement sets</span></span></th>
    <th><span data-ttu-id="b9b24-803"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="b9b24-803"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-804">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-804">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-805">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-805">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b9b24-806">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-806">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b9b24-807">- Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-807">- Selection</span></span><br><span data-ttu-id="b9b24-808">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-808">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-809">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-809">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-810">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-810">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b9b24-811">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-811">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b9b24-812">- Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-812">- Selection</span></span><br><span data-ttu-id="b9b24-813">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-813">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b9b24-814">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="b9b24-814">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b9b24-815">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="b9b24-815">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b9b24-816">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b9b24-816">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b9b24-817">- Selection</span><span class="sxs-lookup"><span data-stu-id="b9b24-817">- Selection</span></span><br><span data-ttu-id="b9b24-818">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b9b24-818">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="b9b24-819">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b9b24-819">See also</span></span>

- [<span data-ttu-id="b9b24-820">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="b9b24-820">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="b9b24-821">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="b9b24-821">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="b9b24-822">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="b9b24-822">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="b9b24-823">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="b9b24-823">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)

---
title: Office 外接程序主机和平台可用性
description: Excel、Word、Outlook、PowerPoint、OneNote 和项目支持的要求集。
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: fe5b1d1278d2c14192fb6fd212f24bb08571d35d
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691123"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="72c1f-103">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="72c1f-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="72c1f-p101">若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。</span><span class="sxs-lookup"><span data-stu-id="72c1f-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="72c1f-p102">通过 MSI 安装的 Office 2016 的生成号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="72c1f-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>
>
> <span data-ttu-id="72c1f-108">Office 2019 的一次性购买的内部版本号是 16.0.10827.20150。</span><span class="sxs-lookup"><span data-stu-id="72c1f-108">The build number for a one-time purchase of Office 2019 is 16.0.10827.20150.</span></span>

## <a name="excel"></a><span data-ttu-id="72c1f-109">Excel</span><span class="sxs-lookup"><span data-stu-id="72c1f-109">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="72c1f-110">平台</span><span class="sxs-lookup"><span data-stu-id="72c1f-110">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="72c1f-111">扩展点</span><span class="sxs-lookup"><span data-stu-id="72c1f-111">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="72c1f-112">API 要求集</span><span class="sxs-lookup"><span data-stu-id="72c1f-112">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="72c1f-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72c1f-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="72c1f-114">Office Online</span></span></td>
    <td> <span data-ttu-id="72c1f-115">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-115">- TaskPane</span></span><br><span data-ttu-id="72c1f-116">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-116">
        - Content</span></span><br><span data-ttu-id="72c1f-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="72c1f-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="72c1f-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72c1f-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72c1f-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72c1f-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72c1f-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72c1f-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72c1f-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="72c1f-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72c1f-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="72c1f-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-127">
        - BindingEvents</span></span><br><span data-ttu-id="72c1f-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-128">
        - CompressedFile</span></span><br><span data-ttu-id="72c1f-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-129">
        - DocumentEvents</span></span><br><span data-ttu-id="72c1f-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-130">
        - File</span></span><br><span data-ttu-id="72c1f-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-131">
        - MatrixBindings</span></span><br><span data-ttu-id="72c1f-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-132">
        - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-133">
        - Selection</span></span><br><span data-ttu-id="72c1f-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-134">
        - Settings</span></span><br><span data-ttu-id="72c1f-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-135">
        - TableBindings</span></span><br><span data-ttu-id="72c1f-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-136">
        - TableCoercion</span></span><br><span data-ttu-id="72c1f-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-137">
        - TextBindings</span></span><br><span data-ttu-id="72c1f-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-138">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-139">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-139">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-140">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-140">- TaskPane</span></span><br><span data-ttu-id="72c1f-141">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-141">
        - Content</span></span><br><span data-ttu-id="72c1f-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </span><span class="sxs-lookup"><span data-stu-id="72c1f-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="72c1f-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72c1f-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72c1f-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72c1f-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72c1f-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72c1f-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72c1f-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="72c1f-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72c1f-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="72c1f-152">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-152">
        - BindingEvents</span></span><br><span data-ttu-id="72c1f-153">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-153">
        - CompressedFile</span></span><br><span data-ttu-id="72c1f-154">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-154">
        - DocumentEvents</span></span><br><span data-ttu-id="72c1f-155">
        - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-155">
        - File</span></span><br><span data-ttu-id="72c1f-156">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-156">
        - MatrixBindings</span></span><br><span data-ttu-id="72c1f-157">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-157">
        - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-158">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-158">
        - Selection</span></span><br><span data-ttu-id="72c1f-159">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-159">
        - Settings</span></span><br><span data-ttu-id="72c1f-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-160">
        - TableBindings</span></span><br><span data-ttu-id="72c1f-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-161">
        - TableCoercion</span></span><br><span data-ttu-id="72c1f-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-162">
        - TextBindings</span></span><br><span data-ttu-id="72c1f-163">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-163">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-164">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-164">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="72c1f-165">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-165">- TaskPane</span></span><br><span data-ttu-id="72c1f-166">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-166">
        - Content</span></span><br><span data-ttu-id="72c1f-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="72c1f-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72c1f-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72c1f-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72c1f-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72c1f-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72c1f-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72c1f-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="72c1f-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72c1f-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="72c1f-177">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-177">- BindingEvents</span></span><br><span data-ttu-id="72c1f-178">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-178">
        - CompressedFile</span></span><br><span data-ttu-id="72c1f-179">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-179">
        - DocumentEvents</span></span><br><span data-ttu-id="72c1f-180">
        - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-180">
        - File</span></span><br><span data-ttu-id="72c1f-181">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-181">
        - ImageCoercion</span></span><br><span data-ttu-id="72c1f-182">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-182">
        - MatrixBindings</span></span><br><span data-ttu-id="72c1f-183">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-183">
        - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-184">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-184">
        - Selection</span></span><br><span data-ttu-id="72c1f-185">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-185">
        - Settings</span></span><br><span data-ttu-id="72c1f-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-186">
        - TableBindings</span></span><br><span data-ttu-id="72c1f-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-187">
        - TableCoercion</span></span><br><span data-ttu-id="72c1f-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-188">
        - TextBindings</span></span><br><span data-ttu-id="72c1f-189">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-189">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-190">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-190">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="72c1f-191">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-191">- TaskPane</span></span><br><span data-ttu-id="72c1f-192">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-192">
        - Content</span></span></td>
    <td><span data-ttu-id="72c1f-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72c1f-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="72c1f-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="72c1f-195">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-195">- BindingEvents</span></span><br><span data-ttu-id="72c1f-196">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-196">
        - CompressedFile</span></span><br><span data-ttu-id="72c1f-197">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-197">
        - DocumentEvents</span></span><br><span data-ttu-id="72c1f-198">
        - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-198">
        - File</span></span><br><span data-ttu-id="72c1f-199">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-199">
        - ImageCoercion</span></span><br><span data-ttu-id="72c1f-200">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-200">
        - MatrixBindings</span></span><br><span data-ttu-id="72c1f-201">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-201">
        - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-202">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-202">
        - Selection</span></span><br><span data-ttu-id="72c1f-203">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-203">
        - Settings</span></span><br><span data-ttu-id="72c1f-204">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-204">
        - TableBindings</span></span><br><span data-ttu-id="72c1f-205">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-205">
        - TableCoercion</span></span><br><span data-ttu-id="72c1f-206">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-206">
        - TextBindings</span></span><br><span data-ttu-id="72c1f-207">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-207">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-208">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-208">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="72c1f-209">
        - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-209">
        - TaskPane</span></span><br><span data-ttu-id="72c1f-210">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-210">
        - Content</span></span></td>
    <td>  <span data-ttu-id="72c1f-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="72c1f-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="72c1f-212">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-212">
        - BindingEvents</span></span><br><span data-ttu-id="72c1f-213">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-213">
        - CompressedFile</span></span><br><span data-ttu-id="72c1f-214">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-214">
        - DocumentEvents</span></span><br><span data-ttu-id="72c1f-215">
        - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-215">
        - File</span></span><br><span data-ttu-id="72c1f-216">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-216">
        - ImageCoercion</span></span><br><span data-ttu-id="72c1f-217">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-217">
        - MatrixBindings</span></span><br><span data-ttu-id="72c1f-218">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-218">
        - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-219">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-219">
        - Selection</span></span><br><span data-ttu-id="72c1f-220">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-220">
        - Settings</span></span><br><span data-ttu-id="72c1f-221">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-221">
        - TableBindings</span></span><br><span data-ttu-id="72c1f-222">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-222">
        - TableCoercion</span></span><br><span data-ttu-id="72c1f-223">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-223">
        - TextBindings</span></span><br><span data-ttu-id="72c1f-224">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-224">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-225">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="72c1f-225">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="72c1f-226">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-226">- TaskPane</span></span><br><span data-ttu-id="72c1f-227">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-227">
        - Content</span></span></td>
    <td><span data-ttu-id="72c1f-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72c1f-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72c1f-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72c1f-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72c1f-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72c1f-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72c1f-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="72c1f-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72c1f-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="72c1f-237">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-237">- BindingEvents</span></span><br><span data-ttu-id="72c1f-238">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-238">
        - CompressedFile</span></span><br><span data-ttu-id="72c1f-239">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-239">
        - DocumentEvents</span></span><br><span data-ttu-id="72c1f-240">
        - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-240">
        - File</span></span><br><span data-ttu-id="72c1f-241">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-241">
        - ImageCoercion</span></span><br><span data-ttu-id="72c1f-242">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-242">
        - MatrixBindings</span></span><br><span data-ttu-id="72c1f-243">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-243">
        - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-244">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-244">
        - Selection</span></span><br><span data-ttu-id="72c1f-245">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-245">
        - Settings</span></span><br><span data-ttu-id="72c1f-246">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-246">
        - TableBindings</span></span><br><span data-ttu-id="72c1f-247">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-247">
        - TableCoercion</span></span><br><span data-ttu-id="72c1f-248">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-248">
        - TextBindings</span></span><br><span data-ttu-id="72c1f-249">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-249">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-250">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="72c1f-250">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="72c1f-251">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-251">- TaskPane</span></span><br><span data-ttu-id="72c1f-252">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-252">
        - Content</span></span><br><span data-ttu-id="72c1f-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="72c1f-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72c1f-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72c1f-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72c1f-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72c1f-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72c1f-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72c1f-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="72c1f-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72c1f-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="72c1f-263">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-263">- BindingEvents</span></span><br><span data-ttu-id="72c1f-264">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-264">
        - CompressedFile</span></span><br><span data-ttu-id="72c1f-265">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-265">
        - DocumentEvents</span></span><br><span data-ttu-id="72c1f-266">
        - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-266">
        - File</span></span><br><span data-ttu-id="72c1f-267">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-267">
        - ImageCoercion</span></span><br><span data-ttu-id="72c1f-268">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-268">
        - MatrixBindings</span></span><br><span data-ttu-id="72c1f-269">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-269">
        - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-270">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-270">
        - PdfFile</span></span><br><span data-ttu-id="72c1f-271">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-271">
        - Selection</span></span><br><span data-ttu-id="72c1f-272">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-272">
        - Settings</span></span><br><span data-ttu-id="72c1f-273">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-273">
        - TableBindings</span></span><br><span data-ttu-id="72c1f-274">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-274">
        - TableCoercion</span></span><br><span data-ttu-id="72c1f-275">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-275">
        - TextBindings</span></span><br><span data-ttu-id="72c1f-276">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-276">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-277">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="72c1f-277">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="72c1f-278">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-278">- TaskPane</span></span><br><span data-ttu-id="72c1f-279">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-279">
        - Content</span></span><br><span data-ttu-id="72c1f-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="72c1f-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72c1f-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72c1f-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72c1f-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72c1f-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72c1f-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72c1f-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="72c1f-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72c1f-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="72c1f-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-290">- BindingEvents</span></span><br><span data-ttu-id="72c1f-291">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-291">
        - CompressedFile</span></span><br><span data-ttu-id="72c1f-292">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-292">
        - DocumentEvents</span></span><br><span data-ttu-id="72c1f-293">
        - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-293">
        - File</span></span><br><span data-ttu-id="72c1f-294">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-294">
        - ImageCoercion</span></span><br><span data-ttu-id="72c1f-295">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-295">
        - MatrixBindings</span></span><br><span data-ttu-id="72c1f-296">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-296">
        - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-297">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-297">
        - PdfFile</span></span><br><span data-ttu-id="72c1f-298">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-298">
        - Selection</span></span><br><span data-ttu-id="72c1f-299">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-299">
        - Settings</span></span><br><span data-ttu-id="72c1f-300">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-300">
        - TableBindings</span></span><br><span data-ttu-id="72c1f-301">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-301">
        - TableCoercion</span></span><br><span data-ttu-id="72c1f-302">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-302">
        - TextBindings</span></span><br><span data-ttu-id="72c1f-303">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-303">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-304">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="72c1f-304">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="72c1f-305">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-305">- TaskPane</span></span><br><span data-ttu-id="72c1f-306">
        - 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-306">
        - Content</span></span></td>
    <td><span data-ttu-id="72c1f-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72c1f-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="72c1f-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="72c1f-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-309">- BindingEvents</span></span><br><span data-ttu-id="72c1f-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-310">
        - CompressedFile</span></span><br><span data-ttu-id="72c1f-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-311">
        - DocumentEvents</span></span><br><span data-ttu-id="72c1f-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-312">
        - File</span></span><br><span data-ttu-id="72c1f-313">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-313">
        - ImageCoercion</span></span><br><span data-ttu-id="72c1f-314">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-314">
        - MatrixBindings</span></span><br><span data-ttu-id="72c1f-315">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-315">
        - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-316">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-316">
        - PdfFile</span></span><br><span data-ttu-id="72c1f-317">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-317">
        - Selection</span></span><br><span data-ttu-id="72c1f-318">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-318">
        - Settings</span></span><br><span data-ttu-id="72c1f-319">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-319">
        - TableBindings</span></span><br><span data-ttu-id="72c1f-320">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-320">
        - TableCoercion</span></span><br><span data-ttu-id="72c1f-321">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-321">
        - TextBindings</span></span><br><span data-ttu-id="72c1f-322">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-322">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="72c1f-323">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="72c1f-323">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="72c1f-324">Outlook</span><span class="sxs-lookup"><span data-stu-id="72c1f-324">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="72c1f-325">平台</span><span class="sxs-lookup"><span data-stu-id="72c1f-325">Platform</span></span></th>
    <th><span data-ttu-id="72c1f-326">扩展点</span><span class="sxs-lookup"><span data-stu-id="72c1f-326">Extension points</span></span></th>
    <th><span data-ttu-id="72c1f-327">API 要求集</span><span class="sxs-lookup"><span data-stu-id="72c1f-327">API requirement sets</span></span></th>
    <th><span data-ttu-id="72c1f-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72c1f-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-329">Office Online</span><span class="sxs-lookup"><span data-stu-id="72c1f-329">Office Online</span></span></td>
    <td> <span data-ttu-id="72c1f-330">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72c1f-330">- Mail Read</span></span><br><span data-ttu-id="72c1f-331">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="72c1f-331">
      - Mail Compose</span></span><br><span data-ttu-id="72c1f-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72c1f-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72c1f-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72c1f-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72c1f-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72c1f-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="72c1f-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="72c1f-340">不可用</span><span class="sxs-lookup"><span data-stu-id="72c1f-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-341">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-341">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-342">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72c1f-342">- Mail Read</span></span><br><span data-ttu-id="72c1f-343">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="72c1f-343">
      - Mail Compose</span></span><br><span data-ttu-id="72c1f-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="72c1f-345">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="72c1f-345">
      - Modules</span></span></td>
    <td> <span data-ttu-id="72c1f-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72c1f-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72c1f-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72c1f-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72c1f-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72c1f-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="72c1f-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="72c1f-353">不可用</span><span class="sxs-lookup"><span data-stu-id="72c1f-353">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-354">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-354">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-355">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72c1f-355">- Mail Read</span></span><br><span data-ttu-id="72c1f-356">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="72c1f-356">
      - Mail Compose</span></span><br><span data-ttu-id="72c1f-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="72c1f-358">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="72c1f-358">
      - Modules</span></span></td>
    <td> <span data-ttu-id="72c1f-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72c1f-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72c1f-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72c1f-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72c1f-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72c1f-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="72c1f-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="72c1f-366">不可用</span><span class="sxs-lookup"><span data-stu-id="72c1f-366">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-367">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-367">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-368">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72c1f-368">- Mail Read</span></span><br><span data-ttu-id="72c1f-369">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="72c1f-369">
      - Mail Compose</span></span><br><span data-ttu-id="72c1f-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="72c1f-371">
      - 模块</span><span class="sxs-lookup"><span data-stu-id="72c1f-371">
      - Modules</span></span></td>
    <td> <span data-ttu-id="72c1f-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72c1f-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72c1f-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72c1f-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="72c1f-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="72c1f-376">暂无</span><span class="sxs-lookup"><span data-stu-id="72c1f-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-377">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-377">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-378">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72c1f-378">- Mail Read</span></span><br><span data-ttu-id="72c1f-379">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="72c1f-379">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="72c1f-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72c1f-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72c1f-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="72c1f-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72c1f-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="72c1f-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="72c1f-384">暂无</span><span class="sxs-lookup"><span data-stu-id="72c1f-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-385">Office 365 for iOS</span><span class="sxs-lookup"><span data-stu-id="72c1f-385">See the Office 365 SDK for iOS.</span></span></td>
    <td> <span data-ttu-id="72c1f-386">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72c1f-386">- Mail Read</span></span><br><span data-ttu-id="72c1f-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72c1f-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72c1f-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72c1f-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72c1f-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="72c1f-393">不可用</span><span class="sxs-lookup"><span data-stu-id="72c1f-393">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-394">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="72c1f-394">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="72c1f-395">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72c1f-395">- Mail Read</span></span><br><span data-ttu-id="72c1f-396">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="72c1f-396">
      - Mail Compose</span></span><br><span data-ttu-id="72c1f-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72c1f-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72c1f-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72c1f-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72c1f-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72c1f-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="72c1f-404">不可用</span><span class="sxs-lookup"><span data-stu-id="72c1f-404">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-405">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="72c1f-405">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="72c1f-406">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72c1f-406">- Mail Read</span></span><br><span data-ttu-id="72c1f-407">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="72c1f-407">
      - Mail Compose</span></span><br><span data-ttu-id="72c1f-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72c1f-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72c1f-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72c1f-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72c1f-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72c1f-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="72c1f-415">不可用</span><span class="sxs-lookup"><span data-stu-id="72c1f-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-416">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="72c1f-416">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="72c1f-417">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72c1f-417">- Mail Read</span></span><br><span data-ttu-id="72c1f-418">
      - 邮件撰写</span><span class="sxs-lookup"><span data-stu-id="72c1f-418">
      - Mail Compose</span></span><br><span data-ttu-id="72c1f-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72c1f-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72c1f-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72c1f-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72c1f-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72c1f-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="72c1f-426">不可用</span><span class="sxs-lookup"><span data-stu-id="72c1f-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-427">Office 365 for Android</span><span class="sxs-lookup"><span data-stu-id="72c1f-427">See the Office 365 SDK for Android.</span></span></td>
    <td> <span data-ttu-id="72c1f-428">- 邮件阅读</span><span class="sxs-lookup"><span data-stu-id="72c1f-428">- Mail Read</span></span><br><span data-ttu-id="72c1f-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72c1f-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72c1f-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72c1f-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72c1f-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="72c1f-435">不可用</span><span class="sxs-lookup"><span data-stu-id="72c1f-435">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="72c1f-436">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="72c1f-436">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="72c1f-437">Word</span><span class="sxs-lookup"><span data-stu-id="72c1f-437">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="72c1f-438">平台</span><span class="sxs-lookup"><span data-stu-id="72c1f-438">Platform</span></span></th>
    <th><span data-ttu-id="72c1f-439">扩展点</span><span class="sxs-lookup"><span data-stu-id="72c1f-439">Extension points</span></span></th>
    <th><span data-ttu-id="72c1f-440">API 要求集</span><span class="sxs-lookup"><span data-stu-id="72c1f-440">API requirement sets</span></span></th>
    <th><span data-ttu-id="72c1f-441"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72c1f-441"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-442">Office Online</span><span class="sxs-lookup"><span data-stu-id="72c1f-442">Office Online</span></span></td>
    <td> <span data-ttu-id="72c1f-443">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-443">- TaskPane</span></span><br><span data-ttu-id="72c1f-444">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-444">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-445">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-445">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="72c1f-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="72c1f-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="72c1f-448">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-448">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72c1f-449">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-449">- BindingEvents</span></span><br><span data-ttu-id="72c1f-450">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72c1f-450">
         - CustomXmlParts</span></span><br><span data-ttu-id="72c1f-451">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-451">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-452">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-452">
         - File</span></span><br><span data-ttu-id="72c1f-453">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-453">
         - HtmlCoercion</span></span><br><span data-ttu-id="72c1f-454">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-454">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-455">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-455">
         - MatrixBindings</span></span><br><span data-ttu-id="72c1f-456">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-456">
         - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-457">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-457">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72c1f-458">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-458">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-459">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-459">
         - Selection</span></span><br><span data-ttu-id="72c1f-460">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-460">
         - Settings</span></span><br><span data-ttu-id="72c1f-461">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-461">
         - TableBindings</span></span><br><span data-ttu-id="72c1f-462">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-462">
         - TableCoercion</span></span><br><span data-ttu-id="72c1f-463">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-463">
         - TextBindings</span></span><br><span data-ttu-id="72c1f-464">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-464">
         - TextCoercion</span></span><br><span data-ttu-id="72c1f-465">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-465">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-466">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-466">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-467">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-467">- TaskPane</span></span><br><span data-ttu-id="72c1f-468">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-468">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-469">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-469">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="72c1f-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="72c1f-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="72c1f-472">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-472">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72c1f-473">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-473">- BindingEvents</span></span><br><span data-ttu-id="72c1f-474">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-474">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-475">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72c1f-475">
         - CustomXmlParts</span></span><br><span data-ttu-id="72c1f-476">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-476">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-477">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-477">
         - File</span></span><br><span data-ttu-id="72c1f-478">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-478">
         - HtmlCoercion</span></span><br><span data-ttu-id="72c1f-479">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-479">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-480">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-480">
         - MatrixBindings</span></span><br><span data-ttu-id="72c1f-481">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-481">
         - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-482">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-482">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72c1f-483">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-483">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-484">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-484">
         - Selection</span></span><br><span data-ttu-id="72c1f-485">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-485">
         - Settings</span></span><br><span data-ttu-id="72c1f-486">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-486">
         - TableBindings</span></span><br><span data-ttu-id="72c1f-487">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-487">
         - TableCoercion</span></span><br><span data-ttu-id="72c1f-488">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-488">
         - TextBindings</span></span><br><span data-ttu-id="72c1f-489">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-489">
         - TextCoercion</span></span><br><span data-ttu-id="72c1f-490">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-490">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-491">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-491">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-492">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-492">- TaskPane</span></span><br><span data-ttu-id="72c1f-493">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-493">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-494">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-494">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="72c1f-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="72c1f-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="72c1f-497">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-497">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72c1f-498">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-498">- BindingEvents</span></span><br><span data-ttu-id="72c1f-499">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-499">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-500">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72c1f-500">
         - CustomXmlParts</span></span><br><span data-ttu-id="72c1f-501">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-501">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-502">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-502">
         - File</span></span><br><span data-ttu-id="72c1f-503">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-503">
         - HtmlCoercion</span></span><br><span data-ttu-id="72c1f-504">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-504">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-505">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-505">
         - MatrixBindings</span></span><br><span data-ttu-id="72c1f-506">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-506">
         - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-507">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-507">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72c1f-508">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-508">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-509">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-509">
         - Selection</span></span><br><span data-ttu-id="72c1f-510">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-510">
         - Settings</span></span><br><span data-ttu-id="72c1f-511">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-511">
         - TableBindings</span></span><br><span data-ttu-id="72c1f-512">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-512">
         - TableCoercion</span></span><br><span data-ttu-id="72c1f-513">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-513">
         - TextBindings</span></span><br><span data-ttu-id="72c1f-514">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-514">
         - TextCoercion</span></span><br><span data-ttu-id="72c1f-515">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-515">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-516">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-516">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-517">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-517">- TaskPane</span></span></td>
    <td> <span data-ttu-id="72c1f-518">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-518">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="72c1f-519">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="72c1f-519">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="72c1f-520">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-520">- BindingEvents</span></span><br><span data-ttu-id="72c1f-521">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-521">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-522">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72c1f-522">
         - CustomXmlParts</span></span><br><span data-ttu-id="72c1f-523">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-523">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-524">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-524">
         - File</span></span><br><span data-ttu-id="72c1f-525">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-525">
         - HtmlCoercion</span></span><br><span data-ttu-id="72c1f-526">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-526">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-527">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-527">
         - MatrixBindings</span></span><br><span data-ttu-id="72c1f-528">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-528">
         - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-529">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-529">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72c1f-530">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-530">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-531">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-531">
         - Selection</span></span><br><span data-ttu-id="72c1f-532">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-532">
         - Settings</span></span><br><span data-ttu-id="72c1f-533">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-533">
         - TableBindings</span></span><br><span data-ttu-id="72c1f-534">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-534">
         - TableCoercion</span></span><br><span data-ttu-id="72c1f-535">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-535">
         - TextBindings</span></span><br><span data-ttu-id="72c1f-536">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-536">
         - TextCoercion</span></span><br><span data-ttu-id="72c1f-537">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-537">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-538">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-538">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-539">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-539">- TaskPane</span></span></td>
    <td> <span data-ttu-id="72c1f-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="72c1f-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="72c1f-541">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-541">- BindingEvents</span></span><br><span data-ttu-id="72c1f-542">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-542">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-543">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72c1f-543">
         - CustomXmlParts</span></span><br><span data-ttu-id="72c1f-544">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-544">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-545">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-545">
         - File</span></span><br><span data-ttu-id="72c1f-546">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-546">
         - HtmlCoercion</span></span><br><span data-ttu-id="72c1f-547">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-547">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-548">
         - MatrixBindings</span></span><br><span data-ttu-id="72c1f-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72c1f-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-551">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-552">
         - Selection</span></span><br><span data-ttu-id="72c1f-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-553">
         - Settings</span></span><br><span data-ttu-id="72c1f-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-554">
         - TableBindings</span></span><br><span data-ttu-id="72c1f-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-555">
         - TableCoercion</span></span><br><span data-ttu-id="72c1f-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-556">
         - TextBindings</span></span><br><span data-ttu-id="72c1f-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-557">
         - TextCoercion</span></span><br><span data-ttu-id="72c1f-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-558">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-559">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="72c1f-559">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="72c1f-560">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-560">- TaskPane</span></span></td>
    <td> <span data-ttu-id="72c1f-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="72c1f-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="72c1f-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="72c1f-564">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="72c1f-564">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="72c1f-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-565">- BindingEvents</span></span><br><span data-ttu-id="72c1f-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-566">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72c1f-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="72c1f-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-568">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-569">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-569">
         - File</span></span><br><span data-ttu-id="72c1f-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="72c1f-571">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-571">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-572">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-572">
         - MatrixBindings</span></span><br><span data-ttu-id="72c1f-573">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-573">
         - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-574">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-574">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72c1f-575">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-575">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-576">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-576">
         - Selection</span></span><br><span data-ttu-id="72c1f-577">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-577">
         - Settings</span></span><br><span data-ttu-id="72c1f-578">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-578">
         - TableBindings</span></span><br><span data-ttu-id="72c1f-579">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-579">
         - TableCoercion</span></span><br><span data-ttu-id="72c1f-580">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-580">
         - TextBindings</span></span><br><span data-ttu-id="72c1f-581">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-581">
         - TextCoercion</span></span><br><span data-ttu-id="72c1f-582">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-582">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-583">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="72c1f-583">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="72c1f-584">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-584">- TaskPane</span></span><br><span data-ttu-id="72c1f-585">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-585">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-586">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-586">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="72c1f-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="72c1f-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="72c1f-589">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="72c1f-589">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="72c1f-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-590">- BindingEvents</span></span><br><span data-ttu-id="72c1f-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-591">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72c1f-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="72c1f-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-593">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-594">
         - File</span></span><br><span data-ttu-id="72c1f-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="72c1f-596">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-596">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-597">
         - MatrixBindings</span></span><br><span data-ttu-id="72c1f-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72c1f-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-600">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-601">
         - Selection</span></span><br><span data-ttu-id="72c1f-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-602">
         - Settings</span></span><br><span data-ttu-id="72c1f-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-603">
         - TableBindings</span></span><br><span data-ttu-id="72c1f-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-604">
         - TableCoercion</span></span><br><span data-ttu-id="72c1f-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-605">
         - TextBindings</span></span><br><span data-ttu-id="72c1f-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-606">
         - TextCoercion</span></span><br><span data-ttu-id="72c1f-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-608">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="72c1f-608">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="72c1f-609">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-609">- TaskPane</span></span><br><span data-ttu-id="72c1f-610">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-610">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-611">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-611">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="72c1f-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="72c1f-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="72c1f-614">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="72c1f-614">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="72c1f-615">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-615">- BindingEvents</span></span><br><span data-ttu-id="72c1f-616">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-616">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-617">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72c1f-617">
         - CustomXmlParts</span></span><br><span data-ttu-id="72c1f-618">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-618">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-619">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-619">
         - File</span></span><br><span data-ttu-id="72c1f-620">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-620">
         - HtmlCoercion</span></span><br><span data-ttu-id="72c1f-621">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-621">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-622">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-622">
         - MatrixBindings</span></span><br><span data-ttu-id="72c1f-623">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-623">
         - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-624">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-624">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72c1f-625">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-625">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-626">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-626">
         - Selection</span></span><br><span data-ttu-id="72c1f-627">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-627">
         - Settings</span></span><br><span data-ttu-id="72c1f-628">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-628">
         - TableBindings</span></span><br><span data-ttu-id="72c1f-629">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-629">
         - TableCoercion</span></span><br><span data-ttu-id="72c1f-630">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-630">
         - TextBindings</span></span><br><span data-ttu-id="72c1f-631">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-631">
         - TextCoercion</span></span><br><span data-ttu-id="72c1f-632">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-632">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-633">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="72c1f-633">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="72c1f-634">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-634">- TaskPane</span></span></td>
    <td> <span data-ttu-id="72c1f-635">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-635">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="72c1f-636">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="72c1f-636">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="72c1f-637">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-637">- BindingEvents</span></span><br><span data-ttu-id="72c1f-638">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-638">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-639">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72c1f-639">
         - CustomXmlParts</span></span><br><span data-ttu-id="72c1f-640">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-640">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-641">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-641">
         - File</span></span><br><span data-ttu-id="72c1f-642">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-642">
         - HtmlCoercion</span></span><br><span data-ttu-id="72c1f-643">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-643">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-644">
         - MatrixBindings</span></span><br><span data-ttu-id="72c1f-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="72c1f-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72c1f-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-647">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-648">
         - Selection</span></span><br><span data-ttu-id="72c1f-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-649">
         - Settings</span></span><br><span data-ttu-id="72c1f-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-650">
         - TableBindings</span></span><br><span data-ttu-id="72c1f-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-651">
         - TableCoercion</span></span><br><span data-ttu-id="72c1f-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72c1f-652">
         - TextBindings</span></span><br><span data-ttu-id="72c1f-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-653">
         - TextCoercion</span></span><br><span data-ttu-id="72c1f-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-654">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="72c1f-655">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="72c1f-655">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="72c1f-656">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="72c1f-656">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="72c1f-657">平台</span><span class="sxs-lookup"><span data-stu-id="72c1f-657">Platform</span></span></th>
    <th><span data-ttu-id="72c1f-658">扩展点</span><span class="sxs-lookup"><span data-stu-id="72c1f-658">Extension points</span></span></th>
    <th><span data-ttu-id="72c1f-659">API 要求集</span><span class="sxs-lookup"><span data-stu-id="72c1f-659">API requirement sets</span></span></th>
    <th><span data-ttu-id="72c1f-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72c1f-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-661">Office Online</span><span class="sxs-lookup"><span data-stu-id="72c1f-661">Office Online</span></span></td>
    <td> <span data-ttu-id="72c1f-662">- 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-662">- Content</span></span><br><span data-ttu-id="72c1f-663">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-663">
         - TaskPane</span></span><br><span data-ttu-id="72c1f-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72c1f-666">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72c1f-666">- ActiveView</span></span><br><span data-ttu-id="72c1f-667">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-667">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-668">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-669">
         - File</span></span><br><span data-ttu-id="72c1f-670">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-670">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-671">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-671">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-672">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-672">
         - Selection</span></span><br><span data-ttu-id="72c1f-673">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-673">
         - Settings</span></span><br><span data-ttu-id="72c1f-674">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-674">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-675">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-675">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-676">- 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-676">- Content</span></span><br><span data-ttu-id="72c1f-677">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-677">
         - TaskPane</span></span><br><span data-ttu-id="72c1f-678">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-678">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-679">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-679">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72c1f-680">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72c1f-680">- ActiveView</span></span><br><span data-ttu-id="72c1f-681">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-681">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-682">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-682">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-683">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-683">
         - File</span></span><br><span data-ttu-id="72c1f-684">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-684">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-685">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-685">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-686">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-686">
         - Selection</span></span><br><span data-ttu-id="72c1f-687">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-687">
         - Settings</span></span><br><span data-ttu-id="72c1f-688">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-688">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-689">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-689">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-690">- 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-690">- Content</span></span><br><span data-ttu-id="72c1f-691">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-691">
         - TaskPane</span></span><br><span data-ttu-id="72c1f-692">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-692">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-693">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-693">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72c1f-694">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72c1f-694">- ActiveView</span></span><br><span data-ttu-id="72c1f-695">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-695">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-696">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-696">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-697">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-697">
         - File</span></span><br><span data-ttu-id="72c1f-698">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-698">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-699">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-699">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-700">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-700">
         - Selection</span></span><br><span data-ttu-id="72c1f-701">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-701">
         - Settings</span></span><br><span data-ttu-id="72c1f-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-702">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-703">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-703">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-704">- 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-704">- Content</span></span><br><span data-ttu-id="72c1f-705">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-705">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="72c1f-706">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="72c1f-706">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="72c1f-707">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72c1f-707">- ActiveView</span></span><br><span data-ttu-id="72c1f-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-708">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-709">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-709">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-710">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-710">
         - File</span></span><br><span data-ttu-id="72c1f-711">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-711">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-712">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-712">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-713">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-713">
         - Selection</span></span><br><span data-ttu-id="72c1f-714">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-714">
         - Settings</span></span><br><span data-ttu-id="72c1f-715">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-715">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-716">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-716">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-717">- 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-717">- Content</span></span><br><span data-ttu-id="72c1f-718">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-718">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="72c1f-719">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="72c1f-719">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="72c1f-720">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72c1f-720">- ActiveView</span></span><br><span data-ttu-id="72c1f-721">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-721">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-722">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-722">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-723">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-723">
         - File</span></span><br><span data-ttu-id="72c1f-724">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-724">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-725">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-725">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-726">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-726">
         - Selection</span></span><br><span data-ttu-id="72c1f-727">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-727">
         - Settings</span></span><br><span data-ttu-id="72c1f-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-728">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-729">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="72c1f-729">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="72c1f-730">- 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-730">- Content</span></span><br><span data-ttu-id="72c1f-731">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-731">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="72c1f-732">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-732">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="72c1f-733">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72c1f-733">- ActiveView</span></span><br><span data-ttu-id="72c1f-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-734">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-735">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-736">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-736">
         - File</span></span><br><span data-ttu-id="72c1f-737">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-737">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-738">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-738">
         - Selection</span></span><br><span data-ttu-id="72c1f-739">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-739">
         - Settings</span></span><br><span data-ttu-id="72c1f-740">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-740">
         - TextCoercion</span></span><br><span data-ttu-id="72c1f-741">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-741">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-742">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="72c1f-742">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="72c1f-743">- 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-743">- Content</span></span><br><span data-ttu-id="72c1f-744">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-744">
         - TaskPane</span></span><br><span data-ttu-id="72c1f-745">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-745">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-746">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-746">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72c1f-747">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72c1f-747">- ActiveView</span></span><br><span data-ttu-id="72c1f-748">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-748">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-749">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-749">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-750">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-750">
         - File</span></span><br><span data-ttu-id="72c1f-751">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-751">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-752">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-752">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-753">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-753">
         - Selection</span></span><br><span data-ttu-id="72c1f-754">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-754">
         - Settings</span></span><br><span data-ttu-id="72c1f-755">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-755">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-756">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="72c1f-756">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="72c1f-757">- 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-757">- Content</span></span><br><span data-ttu-id="72c1f-758">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-758">
         - TaskPane</span></span><br><span data-ttu-id="72c1f-759">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-759">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-760">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-760">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72c1f-761">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72c1f-761">- ActiveView</span></span><br><span data-ttu-id="72c1f-762">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-762">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-763">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-763">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-764">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-764">
         - File</span></span><br><span data-ttu-id="72c1f-765">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-765">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-766">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-767">
         - Selection</span></span><br><span data-ttu-id="72c1f-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-768">
         - Settings</span></span><br><span data-ttu-id="72c1f-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-770">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="72c1f-770">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="72c1f-771">- 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-771">- Content</span></span><br><span data-ttu-id="72c1f-772">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-772">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="72c1f-773">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="72c1f-773">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="72c1f-774">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72c1f-774">- ActiveView</span></span><br><span data-ttu-id="72c1f-775">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-775">
         - CompressedFile</span></span><br><span data-ttu-id="72c1f-776">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-776">
         - DocumentEvents</span></span><br><span data-ttu-id="72c1f-777">
         - File</span><span class="sxs-lookup"><span data-stu-id="72c1f-777">
         - File</span></span><br><span data-ttu-id="72c1f-778">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-778">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-779">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72c1f-779">
         - PdfFile</span></span><br><span data-ttu-id="72c1f-780">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-780">
         - Selection</span></span><br><span data-ttu-id="72c1f-781">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-781">
         - Settings</span></span><br><span data-ttu-id="72c1f-782">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-782">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="72c1f-783">*&ast; - 已添加发布后更新。*</span><span class="sxs-lookup"><span data-stu-id="72c1f-783">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="72c1f-784">OneNote</span><span class="sxs-lookup"><span data-stu-id="72c1f-784">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="72c1f-785">平台</span><span class="sxs-lookup"><span data-stu-id="72c1f-785">Platform</span></span></th>
    <th><span data-ttu-id="72c1f-786">扩展点</span><span class="sxs-lookup"><span data-stu-id="72c1f-786">Extension points</span></span></th>
    <th><span data-ttu-id="72c1f-787">API 要求集</span><span class="sxs-lookup"><span data-stu-id="72c1f-787">API requirement sets</span></span></th>
    <th><span data-ttu-id="72c1f-788"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72c1f-788"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-789">Office Online</span><span class="sxs-lookup"><span data-stu-id="72c1f-789">Office Online</span></span></td>
    <td> <span data-ttu-id="72c1f-790">- 内容</span><span class="sxs-lookup"><span data-stu-id="72c1f-790">- Content</span></span><br><span data-ttu-id="72c1f-791">
         - 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-791">
         - TaskPane</span></span><br><span data-ttu-id="72c1f-792">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-792">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72c1f-793">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-793">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="72c1f-794">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-794">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72c1f-795">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72c1f-795">- DocumentEvents</span></span><br><span data-ttu-id="72c1f-796">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-796">
         - HtmlCoercion</span></span><br><span data-ttu-id="72c1f-797">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-797">
         - ImageCoercion</span></span><br><span data-ttu-id="72c1f-798">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72c1f-798">
         - Settings</span></span><br><span data-ttu-id="72c1f-799">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-799">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="72c1f-800">项目</span><span class="sxs-lookup"><span data-stu-id="72c1f-800">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="72c1f-801">平台</span><span class="sxs-lookup"><span data-stu-id="72c1f-801">Platform</span></span></th>
    <th><span data-ttu-id="72c1f-802">扩展点</span><span class="sxs-lookup"><span data-stu-id="72c1f-802">Extension points</span></span></th>
    <th><span data-ttu-id="72c1f-803">API 要求集</span><span class="sxs-lookup"><span data-stu-id="72c1f-803">API requirement sets</span></span></th>
    <th><span data-ttu-id="72c1f-804"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72c1f-804"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-805">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-805">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-806">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-806">- TaskPane</span></span></td>
    <td> <span data-ttu-id="72c1f-807">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-807">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72c1f-808">- Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-808">- Selection</span></span><br><span data-ttu-id="72c1f-809">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-809">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-810">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-810">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-811">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-811">- TaskPane</span></span></td>
    <td> <span data-ttu-id="72c1f-812">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-812">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72c1f-813">- Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-813">- Selection</span></span><br><span data-ttu-id="72c1f-814">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-814">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72c1f-815">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="72c1f-815">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="72c1f-816">- 任务窗格</span><span class="sxs-lookup"><span data-stu-id="72c1f-816">- TaskPane</span></span></td>
    <td> <span data-ttu-id="72c1f-817">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72c1f-817">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72c1f-818">- Selection</span><span class="sxs-lookup"><span data-stu-id="72c1f-818">- Selection</span></span><br><span data-ttu-id="72c1f-819">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72c1f-819">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="72c1f-820">另请参阅</span><span class="sxs-lookup"><span data-stu-id="72c1f-820">See also</span></span>

- [<span data-ttu-id="72c1f-821">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="72c1f-821">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="72c1f-822">通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="72c1f-822">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="72c1f-823">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="72c1f-823">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="72c1f-824">适用于 Office 的 JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="72c1f-824">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)

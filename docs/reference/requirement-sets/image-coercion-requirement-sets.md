---
title: 图像强制要求集
description: 支持跨 Excel、PowerPoint 和 Word 的 Office 外接程序对图像强制要求集的支持。
ms.date: 07/11/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 046a3f1f16d8b48cddbd64bddf80a31ed1e50583
ms.sourcegitcommit: 61f8f02193ce05da957418d938f0d94cb12c468d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/11/2019
ms.locfileid: "35633989"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="69960-103">图像强制要求集</span><span class="sxs-lookup"><span data-stu-id="69960-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="69960-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="69960-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="69960-107">Office 外接程序在多个 Office 版本中运行。</span><span class="sxs-lookup"><span data-stu-id="69960-107">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="69960-108">下表列出了图像强制要求集、支持该要求集的 Office 主机应用程序, 以及 Office 应用程序的内部版本号或版本号。</span><span class="sxs-lookup"><span data-stu-id="69960-108">The following table lists the Image Coercion requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="69960-109">ImageCoercion 1。1</span><span class="sxs-lookup"><span data-stu-id="69960-109">ImageCoercion 1.1</span></span>

<span data-ttu-id="69960-110">在使用[`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-)方法写入数据时, ImageCoercion`Office.CoercionType.Image`1.1 支持转换为 image ()。</span><span class="sxs-lookup"><span data-stu-id="69960-110">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="69960-111">支持以下主机:</span><span class="sxs-lookup"><span data-stu-id="69960-111">The following hosts are supported:</span></span>

- <span data-ttu-id="69960-112">Excel 2013 及更高版本的 Windows</span><span class="sxs-lookup"><span data-stu-id="69960-112">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="69960-113">Excel 2016 及更高版本 Mac</span><span class="sxs-lookup"><span data-stu-id="69960-113">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="69960-114">在 web 上的 Excel</span><span class="sxs-lookup"><span data-stu-id="69960-114">Excel on the web</span></span>
- <span data-ttu-id="69960-115">IPad 上的 Excel</span><span class="sxs-lookup"><span data-stu-id="69960-115">Excel on iPad</span></span>
- <span data-ttu-id="69960-116">在 web 上的 OneNote</span><span class="sxs-lookup"><span data-stu-id="69960-116">OneNote on the web</span></span>
- <span data-ttu-id="69960-117">PowerPoint 2013 及更高版本 Windows</span><span class="sxs-lookup"><span data-stu-id="69960-117">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="69960-118">PowerPoint 2016 及更高版本 Mac</span><span class="sxs-lookup"><span data-stu-id="69960-118">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="69960-119">在 web 上的 PowerPoint</span><span class="sxs-lookup"><span data-stu-id="69960-119">PowerPoint on the web</span></span>
- <span data-ttu-id="69960-120">IPad 上的 PowerPoint</span><span class="sxs-lookup"><span data-stu-id="69960-120">PowerPoint on iPad</span></span>
- <span data-ttu-id="69960-121">Word 2013 及更高版本的 Windows</span><span class="sxs-lookup"><span data-stu-id="69960-121">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="69960-122">Word 2016 及更高版本 Mac</span><span class="sxs-lookup"><span data-stu-id="69960-122">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="69960-123">在 web 上的 Word</span><span class="sxs-lookup"><span data-stu-id="69960-123">Word on the web</span></span>
- <span data-ttu-id="69960-124">iPad 上的 Word</span><span class="sxs-lookup"><span data-stu-id="69960-124">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="69960-125">ImageCoercion 1。2</span><span class="sxs-lookup"><span data-stu-id="69960-125">ImageCoercion 1.2</span></span>

<span data-ttu-id="69960-126">ImageCoercion 1.2 支持在使用`Office.CoercionType.XmlSvg` [`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-)方法写入数据时转换为 SVG 格式 ()。</span><span class="sxs-lookup"><span data-stu-id="69960-126">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="69960-127">支持以下主机:</span><span class="sxs-lookup"><span data-stu-id="69960-127">The following hosts are supported:</span></span>

- <span data-ttu-id="69960-128">Windows 上的 Excel (连接到 Office 365 订阅)</span><span class="sxs-lookup"><span data-stu-id="69960-128">Excel on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="69960-129">Mac 上的 Excel (连接到 Office 365 订阅)</span><span class="sxs-lookup"><span data-stu-id="69960-129">Excel on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="69960-130">在 web 上的 Excel</span><span class="sxs-lookup"><span data-stu-id="69960-130">Excel on the web</span></span>
- <span data-ttu-id="69960-131">Windows 上的 PowerPoint (连接到 Office 365 订阅)</span><span class="sxs-lookup"><span data-stu-id="69960-131">PowerPoint on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="69960-132">PowerPoint on Mac (连接到 Office 365 订阅)</span><span class="sxs-lookup"><span data-stu-id="69960-132">PowerPoint on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="69960-133">在 web 上的 PowerPoint</span><span class="sxs-lookup"><span data-stu-id="69960-133">PowerPoint on the web</span></span>
- <span data-ttu-id="69960-134">Windows 上的 Word (连接到 Office 365 订阅)</span><span class="sxs-lookup"><span data-stu-id="69960-134">Word on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="69960-135">Mac 上的 Word (连接到 Office 365 订阅)</span><span class="sxs-lookup"><span data-stu-id="69960-135">Word on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="69960-136">在 web 上的 Word</span><span class="sxs-lookup"><span data-stu-id="69960-136">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="69960-137">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="69960-137">Office Common API requirement sets</span></span>

<span data-ttu-id="69960-138">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="69960-138">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="69960-139">另请参阅</span><span class="sxs-lookup"><span data-stu-id="69960-139">See also</span></span>

- [<span data-ttu-id="69960-140">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="69960-140">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="69960-141">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="69960-141">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="69960-142">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="69960-142">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)

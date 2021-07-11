---
title: 图像强制要求集
description: 通过跨 Office、PowerPoint 和 Word 的外接程序支持图像强制要求集Excel外接程序。
ms.date: 02/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 29614718378fd51013360a2a922e11f89bca14b8
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350216"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="3dbd2-103">图像强制要求集</span><span class="sxs-lookup"><span data-stu-id="3dbd2-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="3dbd2-p101">要求集是指已命名的 API 成员组。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="3dbd2-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="3dbd2-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="3dbd2-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="3dbd2-108">ImageCoercion 1.1 支持在使用 方法 () `Office.CoercionType.Image` 图像图像 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 转换。</span><span class="sxs-lookup"><span data-stu-id="3dbd2-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="3dbd2-109">支持以下应用程序。</span><span class="sxs-lookup"><span data-stu-id="3dbd2-109">The following applications are supported.</span></span>

- <span data-ttu-id="3dbd2-110">Excel 2013 年 10 月及Windows</span><span class="sxs-lookup"><span data-stu-id="3dbd2-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="3dbd2-111">Excel 2016 Mac 及更高版本</span><span class="sxs-lookup"><span data-stu-id="3dbd2-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="3dbd2-112">iPad 版 Excel</span><span class="sxs-lookup"><span data-stu-id="3dbd2-112">Excel on iPad</span></span>
- <span data-ttu-id="3dbd2-113">OneNote 网页版</span><span class="sxs-lookup"><span data-stu-id="3dbd2-113">OneNote on the web</span></span>
- <span data-ttu-id="3dbd2-114">PowerPoint 2013 及更高版本Windows</span><span class="sxs-lookup"><span data-stu-id="3dbd2-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="3dbd2-115">PowerPoint 2016 Mac 及更高版本</span><span class="sxs-lookup"><span data-stu-id="3dbd2-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="3dbd2-116">PowerPoint 网页版</span><span class="sxs-lookup"><span data-stu-id="3dbd2-116">PowerPoint on the web</span></span>
- <span data-ttu-id="3dbd2-117">iPad 版 PowerPoint</span><span class="sxs-lookup"><span data-stu-id="3dbd2-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="3dbd2-118">Windows 版 Word 2013 及更高版本</span><span class="sxs-lookup"><span data-stu-id="3dbd2-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="3dbd2-119">Mac 版 Word 2016 及更高版本</span><span class="sxs-lookup"><span data-stu-id="3dbd2-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="3dbd2-120">Word 网页版</span><span class="sxs-lookup"><span data-stu-id="3dbd2-120">Word on the web</span></span>
- <span data-ttu-id="3dbd2-121">iPad 版 Word</span><span class="sxs-lookup"><span data-stu-id="3dbd2-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="3dbd2-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="3dbd2-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="3dbd2-123">ImageCoercion 1.2 支持在使用 () 写入数据时转换为 SVG `Office.CoercionType.XmlSvg` [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 格式。</span><span class="sxs-lookup"><span data-stu-id="3dbd2-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="3dbd2-124">支持以下应用程序。</span><span class="sxs-lookup"><span data-stu-id="3dbd2-124">The following applications are supported.</span></span>

- <span data-ttu-id="3dbd2-125">Excel连接到Windows (订阅Microsoft 365时) </span><span class="sxs-lookup"><span data-stu-id="3dbd2-125">Excel on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="3dbd2-126">Excel Mac (连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="3dbd2-126">Excel on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="3dbd2-127">PowerPoint连接到Windows (订阅Microsoft 365时) </span><span class="sxs-lookup"><span data-stu-id="3dbd2-127">PowerPoint on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="3dbd2-128">PowerPoint Mac (连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="3dbd2-128">PowerPoint on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="3dbd2-129">PowerPoint 网页版</span><span class="sxs-lookup"><span data-stu-id="3dbd2-129">PowerPoint on the web</span></span>
- <span data-ttu-id="3dbd2-130">Word on Windows (连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="3dbd2-130">Word on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="3dbd2-131">Mac 版 Word (连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="3dbd2-131">Word on Mac (connected to a Microsoft 365 subscription)</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="3dbd2-132">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="3dbd2-132">Office Common API requirement sets</span></span>

<span data-ttu-id="3dbd2-133">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="3dbd2-133">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="3dbd2-134">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3dbd2-134">See also</span></span>

- [<span data-ttu-id="3dbd2-135">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="3dbd2-135">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="3dbd2-136">指定 Office 应用程序和 API 要求集</span><span class="sxs-lookup"><span data-stu-id="3dbd2-136">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="3dbd2-137">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="3dbd2-137">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)

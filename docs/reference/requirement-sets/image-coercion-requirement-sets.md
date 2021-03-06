---
title: 图像强制要求集
description: 支持跨 Excel、PowerPoint 和 Word 使用 Office 加载项的图像强制要求集。
ms.date: 02/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 52ce46a46580500f5a292bf898674d4798378319
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505526"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="6039f-103">图像强制要求集</span><span class="sxs-lookup"><span data-stu-id="6039f-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="6039f-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="6039f-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="6039f-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="6039f-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="6039f-108">ImageCoercion 1.1 支持在 () `Office.CoercionType.Image` 写入数据时转换为图像 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 图像。</span><span class="sxs-lookup"><span data-stu-id="6039f-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="6039f-109">支持以下应用程序：</span><span class="sxs-lookup"><span data-stu-id="6039f-109">The following applications are supported:</span></span>

- <span data-ttu-id="6039f-110">Windows 版 Excel 2013 及更高版本</span><span class="sxs-lookup"><span data-stu-id="6039f-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="6039f-111">Mac 版 Excel 2016 及更高版本</span><span class="sxs-lookup"><span data-stu-id="6039f-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="6039f-112">iPad 版 Excel</span><span class="sxs-lookup"><span data-stu-id="6039f-112">Excel on iPad</span></span>
- <span data-ttu-id="6039f-113">OneNote 网页版</span><span class="sxs-lookup"><span data-stu-id="6039f-113">OneNote on the web</span></span>
- <span data-ttu-id="6039f-114">Windows 版 PowerPoint 2013 及更高版本</span><span class="sxs-lookup"><span data-stu-id="6039f-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="6039f-115">Mac 版 PowerPoint 2016 及更高版本</span><span class="sxs-lookup"><span data-stu-id="6039f-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="6039f-116">PowerPoint 网页版</span><span class="sxs-lookup"><span data-stu-id="6039f-116">PowerPoint on the web</span></span>
- <span data-ttu-id="6039f-117">iPad 版 PowerPoint</span><span class="sxs-lookup"><span data-stu-id="6039f-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="6039f-118">Windows 版 Word 2013 及更高版本</span><span class="sxs-lookup"><span data-stu-id="6039f-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="6039f-119">Mac 版 Word 2016 及更高版本</span><span class="sxs-lookup"><span data-stu-id="6039f-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="6039f-120">Word 网页版</span><span class="sxs-lookup"><span data-stu-id="6039f-120">Word on the web</span></span>
- <span data-ttu-id="6039f-121">iPad 版 Word</span><span class="sxs-lookup"><span data-stu-id="6039f-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="6039f-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="6039f-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="6039f-123">ImageCoercion 1.2 支持在 () 写入数据时转换为 SVG `Office.CoercionType.XmlSvg` [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 格式。</span><span class="sxs-lookup"><span data-stu-id="6039f-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="6039f-124">支持以下应用程序：</span><span class="sxs-lookup"><span data-stu-id="6039f-124">The following applications are supported:</span></span>

- <span data-ttu-id="6039f-125">Windows 版 Excel (Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="6039f-125">Excel on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="6039f-126">Mac 版 Excel (Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="6039f-126">Excel on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="6039f-127">Windows 版 PowerPoint (连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="6039f-127">PowerPoint on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="6039f-128">Mac 版 PowerPoint (Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="6039f-128">PowerPoint on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="6039f-129">PowerPoint 网页版</span><span class="sxs-lookup"><span data-stu-id="6039f-129">PowerPoint on the web</span></span>
- <span data-ttu-id="6039f-130">Windows 版 Word (Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="6039f-130">Word on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="6039f-131">Mac 版 Word (Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="6039f-131">Word on Mac (connected to a Microsoft 365 subscription)</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="6039f-132">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="6039f-132">Office Common API requirement sets</span></span>

<span data-ttu-id="6039f-133">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="6039f-133">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="6039f-134">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6039f-134">See also</span></span>

- [<span data-ttu-id="6039f-135">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="6039f-135">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="6039f-136">指定 Office 应用程序和 API 要求集</span><span class="sxs-lookup"><span data-stu-id="6039f-136">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="6039f-137">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="6039f-137">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)

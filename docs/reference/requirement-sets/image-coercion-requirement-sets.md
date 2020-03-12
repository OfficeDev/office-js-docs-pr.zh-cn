---
title: 图像强制要求集
description: 支持跨 Excel、PowerPoint 和 Word 的 Office 外接程序对图像强制要求集的支持。
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 83817bfc7cf8a193138a805b0e90b4357d605801
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596968"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="cf04b-103">图像强制要求集</span><span class="sxs-lookup"><span data-stu-id="cf04b-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="cf04b-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="cf04b-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="cf04b-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="cf04b-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="cf04b-108">在使用[`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-)方法写入数据时，ImageCoercion`Office.CoercionType.Image`1.1 支持转换为 image （）。</span><span class="sxs-lookup"><span data-stu-id="cf04b-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="cf04b-109">支持以下主机：</span><span class="sxs-lookup"><span data-stu-id="cf04b-109">The following hosts are supported:</span></span>

- <span data-ttu-id="cf04b-110">Excel 2013 及更高版本的 Windows</span><span class="sxs-lookup"><span data-stu-id="cf04b-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="cf04b-111">Excel 2016 及更高版本 Mac</span><span class="sxs-lookup"><span data-stu-id="cf04b-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="cf04b-112">iPad 版 Excel</span><span class="sxs-lookup"><span data-stu-id="cf04b-112">Excel on iPad</span></span>
- <span data-ttu-id="cf04b-113">OneNote 网页版</span><span class="sxs-lookup"><span data-stu-id="cf04b-113">OneNote on the web</span></span>
- <span data-ttu-id="cf04b-114">PowerPoint 2013 及更高版本 Windows</span><span class="sxs-lookup"><span data-stu-id="cf04b-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="cf04b-115">PowerPoint 2016 及更高版本 Mac</span><span class="sxs-lookup"><span data-stu-id="cf04b-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="cf04b-116">PowerPoint 网页版</span><span class="sxs-lookup"><span data-stu-id="cf04b-116">PowerPoint on the web</span></span>
- <span data-ttu-id="cf04b-117">iPad 版 PowerPoint</span><span class="sxs-lookup"><span data-stu-id="cf04b-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="cf04b-118">Windows 版 Word 2013 及更高版本</span><span class="sxs-lookup"><span data-stu-id="cf04b-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="cf04b-119">Mac 版 Word 2016 及更高版本</span><span class="sxs-lookup"><span data-stu-id="cf04b-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="cf04b-120">Word 网页版</span><span class="sxs-lookup"><span data-stu-id="cf04b-120">Word on the web</span></span>
- <span data-ttu-id="cf04b-121">iPad 版 Word</span><span class="sxs-lookup"><span data-stu-id="cf04b-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="cf04b-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="cf04b-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="cf04b-123">ImageCoercion 1.2 支持在使用`Office.CoercionType.XmlSvg` [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-)方法写入数据时转换为 SVG 格式（）。</span><span class="sxs-lookup"><span data-stu-id="cf04b-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="cf04b-124">支持以下主机：</span><span class="sxs-lookup"><span data-stu-id="cf04b-124">The following hosts are supported:</span></span>

- <span data-ttu-id="cf04b-125">Windows 上的 Excel （连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="cf04b-125">Excel on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="cf04b-126">Mac 上的 Excel （连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="cf04b-126">Excel on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="cf04b-127">Windows 上的 PowerPoint （连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="cf04b-127">PowerPoint on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="cf04b-128">PowerPoint on Mac （连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="cf04b-128">PowerPoint on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="cf04b-129">PowerPoint 网页版</span><span class="sxs-lookup"><span data-stu-id="cf04b-129">PowerPoint on the web</span></span>
- <span data-ttu-id="cf04b-130">Windows 上的 Word （连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="cf04b-130">Word on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="cf04b-131">Mac 上的 Word （连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="cf04b-131">Word on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="cf04b-132">Word 网页版</span><span class="sxs-lookup"><span data-stu-id="cf04b-132">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="cf04b-133">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="cf04b-133">Office Common API requirement sets</span></span>

<span data-ttu-id="cf04b-134">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="cf04b-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="cf04b-135">另请参阅</span><span class="sxs-lookup"><span data-stu-id="cf04b-135">See also</span></span>

- [<span data-ttu-id="cf04b-136">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="cf04b-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="cf04b-137">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="cf04b-137">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="cf04b-138">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="cf04b-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)

---
title: 图像强制要求集
description: 支持跨 Excel、PowerPoint 和 Word 的 Office 外接程序对图像强制要求集的支持。
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 7140099757c6e4b5ad405723d5fed95fded6d919
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293546"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="1a21f-103">图像强制要求集</span><span class="sxs-lookup"><span data-stu-id="1a21f-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="1a21f-104">要求集是指各组已命名的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="1a21f-104">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="1a21f-105">Office 外接程序使用清单中指定的要求集或使用运行时检查来确定 Office 应用程序是否支持加载项所需的 Api。</span><span class="sxs-lookup"><span data-stu-id="1a21f-105">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs.</span></span> <span data-ttu-id="1a21f-106">有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="1a21f-106">For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="1a21f-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="1a21f-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="1a21f-108">使用 ImageCoercion 1.1，可以 `Office.CoercionType.Image` 在使用方法写入数据时转换为) 的图像 ([`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 。</span><span class="sxs-lookup"><span data-stu-id="1a21f-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="1a21f-109">支持以下应用程序：</span><span class="sxs-lookup"><span data-stu-id="1a21f-109">The following applications are supported:</span></span>

- <span data-ttu-id="1a21f-110">Excel 2013 及更高版本的 Windows</span><span class="sxs-lookup"><span data-stu-id="1a21f-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="1a21f-111">Excel 2016 及更高版本 Mac</span><span class="sxs-lookup"><span data-stu-id="1a21f-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="1a21f-112">iPad 版 Excel</span><span class="sxs-lookup"><span data-stu-id="1a21f-112">Excel on iPad</span></span>
- <span data-ttu-id="1a21f-113">OneNote 网页版</span><span class="sxs-lookup"><span data-stu-id="1a21f-113">OneNote on the web</span></span>
- <span data-ttu-id="1a21f-114">PowerPoint 2013 及更高版本 Windows</span><span class="sxs-lookup"><span data-stu-id="1a21f-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="1a21f-115">PowerPoint 2016 及更高版本 Mac</span><span class="sxs-lookup"><span data-stu-id="1a21f-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="1a21f-116">PowerPoint 网页版</span><span class="sxs-lookup"><span data-stu-id="1a21f-116">PowerPoint on the web</span></span>
- <span data-ttu-id="1a21f-117">iPad 版 PowerPoint</span><span class="sxs-lookup"><span data-stu-id="1a21f-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="1a21f-118">Windows 版 Word 2013 及更高版本</span><span class="sxs-lookup"><span data-stu-id="1a21f-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="1a21f-119">Mac 版 Word 2016 及更高版本</span><span class="sxs-lookup"><span data-stu-id="1a21f-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="1a21f-120">Word 网页版</span><span class="sxs-lookup"><span data-stu-id="1a21f-120">Word on the web</span></span>
- <span data-ttu-id="1a21f-121">iPad 版 Word</span><span class="sxs-lookup"><span data-stu-id="1a21f-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="1a21f-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="1a21f-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="1a21f-123">ImageCoercion 1.2 支持在 `Office.CoercionType.XmlSvg` 使用方法写入数据时 () 转换为 SVG 格式 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 。</span><span class="sxs-lookup"><span data-stu-id="1a21f-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="1a21f-124">支持以下应用程序：</span><span class="sxs-lookup"><span data-stu-id="1a21f-124">The following applications are supported:</span></span>

- <span data-ttu-id="1a21f-125">连接到 Microsoft 365 订阅的 Windows (上的 Excel) </span><span class="sxs-lookup"><span data-stu-id="1a21f-125">Excel on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="1a21f-126">连接到 Microsoft 365 订阅的 Mac 上的 Excel () </span><span class="sxs-lookup"><span data-stu-id="1a21f-126">Excel on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="1a21f-127">连接到 Microsoft 365 订阅的 Windows (上的 PowerPoint) </span><span class="sxs-lookup"><span data-stu-id="1a21f-127">PowerPoint on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="1a21f-128">连接到 Microsoft 365 订阅的 Mac 版上的 PowerPoint () </span><span class="sxs-lookup"><span data-stu-id="1a21f-128">PowerPoint on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="1a21f-129">PowerPoint 网页版</span><span class="sxs-lookup"><span data-stu-id="1a21f-129">PowerPoint on the web</span></span>
- <span data-ttu-id="1a21f-130">连接到 Microsoft 365 订阅的 Windows (上的 Word) </span><span class="sxs-lookup"><span data-stu-id="1a21f-130">Word on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="1a21f-131">连接到 Microsoft 365 订阅的 Mac 上的 Word () </span><span class="sxs-lookup"><span data-stu-id="1a21f-131">Word on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="1a21f-132">Word 网页版</span><span class="sxs-lookup"><span data-stu-id="1a21f-132">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="1a21f-133">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="1a21f-133">Office Common API requirement sets</span></span>

<span data-ttu-id="1a21f-134">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="1a21f-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="1a21f-135">另请参阅</span><span class="sxs-lookup"><span data-stu-id="1a21f-135">See also</span></span>

- [<span data-ttu-id="1a21f-136">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="1a21f-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="1a21f-137">指定 Office 应用程序和 API 要求</span><span class="sxs-lookup"><span data-stu-id="1a21f-137">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="1a21f-138">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="1a21f-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)

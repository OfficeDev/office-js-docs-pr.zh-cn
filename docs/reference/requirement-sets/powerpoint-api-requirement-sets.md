---
title: PowerPoint JavaScript API 要求集
description: 了解有关 PowerPoint JavaScript API 要求集的详细信息
ms.date: 03/11/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: a82d73087b19fbce12f571a2bad61e866ab62f86
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611328"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="07a3e-103">PowerPoint JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="07a3e-103">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="07a3e-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="07a3e-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="07a3e-107">下表列出了 PowerPoint 要求集、支持这些要求集的 Office 主机应用程序，以及这些应用程序的内部版本或发布日期。</span><span class="sxs-lookup"><span data-stu-id="07a3e-107">The following table lists the PowerPoint requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="07a3e-108">要求集</span><span class="sxs-lookup"><span data-stu-id="07a3e-108">Requirement set</span></span>  |  <span data-ttu-id="07a3e-109">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="07a3e-109">Office on Windows</span></span><br><span data-ttu-id="07a3e-110">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="07a3e-110">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="07a3e-111">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="07a3e-111">Office on iPad</span></span><br><span data-ttu-id="07a3e-112">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="07a3e-112">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="07a3e-113">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="07a3e-113">Office on Mac</span></span><br><span data-ttu-id="07a3e-114">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="07a3e-114">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="07a3e-115">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="07a3e-115">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="07a3e-116">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="07a3e-116">PowerPointApi 1.1</span></span> | <span data-ttu-id="07a3e-117">版本 1810（内部版本 11001.20074）或更高版本</span><span class="sxs-lookup"><span data-stu-id="07a3e-117">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="07a3e-118">2.17 或更高版本</span><span class="sxs-lookup"><span data-stu-id="07a3e-118">2.17 or later</span></span> | <span data-ttu-id="07a3e-119">16.19 或更高版本</span><span class="sxs-lookup"><span data-stu-id="07a3e-119">16.19 or later</span></span> | <span data-ttu-id="07a3e-120">2018 年 10 月</span><span class="sxs-lookup"><span data-stu-id="07a3e-120">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="07a3e-121">Office 版本和内部版本号</span><span class="sxs-lookup"><span data-stu-id="07a3e-121">Office versions and build numbers</span></span>

<span data-ttu-id="07a3e-122">有关 Office 版本和内部版本号的详细信息，请参阅：</span><span class="sxs-lookup"><span data-stu-id="07a3e-122">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="07a3e-123">PowerPoint JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="07a3e-123">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="07a3e-124">PowerPoint JavaScript API 1.1 包含用于创建新演示文稿的单一 API。</span><span class="sxs-lookup"><span data-stu-id="07a3e-124">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="07a3e-125">有关该 API 的详细信息，请参阅[适用于 PowerPoint 的 JavaScript API](../../powerpoint/powerpoint-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="07a3e-125">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="07a3e-126">运行时要求支持检查</span><span class="sxs-lookup"><span data-stu-id="07a3e-126">Runtime requirement support check</span></span>

<span data-ttu-id="07a3e-127">在运行时，加载项可以执行下列检查，确定特定主机是否支持 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="07a3e-127">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="07a3e-128">基于清单的要求支持检查</span><span class="sxs-lookup"><span data-stu-id="07a3e-128">Manifest-based requirement support check</span></span>

<span data-ttu-id="07a3e-129">使用加载项清单中的 `Requirements` 元素指定加载项必须使用的关键要求集或 API 成员。</span><span class="sxs-lookup"><span data-stu-id="07a3e-129">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="07a3e-130">如果 Office 主机或平台不支持 `Requirements` 元素中指定的要求集或 API 成员，则加载项将无法在该主机或平台上运行，并且不会显示在“我的加载项”中。</span><span class="sxs-lookup"><span data-stu-id="07a3e-130">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="07a3e-131">下面的代码示例展示了加载所有 支持第 1.1 版 OneNoteApi 要求集的 Office 主机应用程序的外接程序。</span><span class="sxs-lookup"><span data-stu-id="07a3e-131">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="07a3e-132">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="07a3e-132">Office Common API requirement sets</span></span>

<span data-ttu-id="07a3e-133">大多数 PowerPoint 加载项功能都来自通用 API 集。</span><span class="sxs-lookup"><span data-stu-id="07a3e-133">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="07a3e-134">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="07a3e-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="07a3e-135">另请参阅</span><span class="sxs-lookup"><span data-stu-id="07a3e-135">See also</span></span>

- [<span data-ttu-id="07a3e-136">PowerPoint JavaScript API 参考文档</span><span class="sxs-lookup"><span data-stu-id="07a3e-136">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="07a3e-137">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="07a3e-137">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="07a3e-138">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="07a3e-138">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="07a3e-139">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="07a3e-139">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)

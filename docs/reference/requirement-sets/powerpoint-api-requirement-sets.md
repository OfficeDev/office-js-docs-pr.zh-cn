---
title: PowerPoint JavaScript API 要求集
description: ''
ms.date: 03/11/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: ef76077c3a2a975fae8a0dc101e8e1b42ef66094
ms.sourcegitcommit: 05b73cdec5f4db7f0b8d48a5a552ee296a0332ca
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42600695"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="328ae-102">PowerPoint JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="328ae-102">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="328ae-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="328ae-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="328ae-106">下表列出了 PowerPoint 要求集、支持这些要求集的 Office 主机应用程序，以及这些应用程序的内部版本或发布日期。</span><span class="sxs-lookup"><span data-stu-id="328ae-106">The following table lists the PowerPoint requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="328ae-107">要求集</span><span class="sxs-lookup"><span data-stu-id="328ae-107">Requirement set</span></span>  |  <span data-ttu-id="328ae-108">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="328ae-108">Office on Windows</span></span><br><span data-ttu-id="328ae-109">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="328ae-109">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="328ae-110">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="328ae-110">Office on iPad</span></span><br><span data-ttu-id="328ae-111">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="328ae-111">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="328ae-112">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="328ae-112">Office on Mac</span></span><br><span data-ttu-id="328ae-113">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="328ae-113">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="328ae-114">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="328ae-114">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="328ae-115">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="328ae-115">PowerPointApi 1.1</span></span> | <span data-ttu-id="328ae-116">版本 1810（内部版本 11001.20074）或更高版本</span><span class="sxs-lookup"><span data-stu-id="328ae-116">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="328ae-117">2.17 或更高版本</span><span class="sxs-lookup"><span data-stu-id="328ae-117">2.17 or later</span></span> | <span data-ttu-id="328ae-118">16.19 或更高版本</span><span class="sxs-lookup"><span data-stu-id="328ae-118">16.19 or later</span></span> | <span data-ttu-id="328ae-119">2018 年 10 月</span><span class="sxs-lookup"><span data-stu-id="328ae-119">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="328ae-120">Office 版本和内部版本号</span><span class="sxs-lookup"><span data-stu-id="328ae-120">Office versions and build numbers</span></span>

<span data-ttu-id="328ae-121">有关 Office 版本和内部版本号的详细信息，请参阅：</span><span class="sxs-lookup"><span data-stu-id="328ae-121">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="328ae-122">PowerPoint JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="328ae-122">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="328ae-123">PowerPoint JavaScript API 1.1 包含用于创建新演示文稿的单一 API。</span><span class="sxs-lookup"><span data-stu-id="328ae-123">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="328ae-124">有关该 API 的详细信息，请参阅[适用于 PowerPoint 的 JavaScript API](../../powerpoint/powerpoint-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="328ae-124">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="328ae-125">运行时要求支持检查</span><span class="sxs-lookup"><span data-stu-id="328ae-125">Runtime requirement support check</span></span>

<span data-ttu-id="328ae-126">在运行时，加载项可以执行下列检查，确定特定主机是否支持 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="328ae-126">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="328ae-127">基于清单的要求支持检查</span><span class="sxs-lookup"><span data-stu-id="328ae-127">Manifest-based requirement support check</span></span>

<span data-ttu-id="328ae-128">使用加载项清单中的 `Requirements` 元素指定加载项必须使用的关键要求集或 API 成员。</span><span class="sxs-lookup"><span data-stu-id="328ae-128">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="328ae-129">如果 Office 主机或平台不支持 `Requirements` 元素中指定的要求集或 API 成员，则加载项将无法在该主机或平台上运行，并且不会显示在“我的加载项”中。</span><span class="sxs-lookup"><span data-stu-id="328ae-129">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="328ae-130">下面的代码示例展示了加载所有 支持第 1.1 版 OneNoteApi 要求集的 Office 主机应用程序的外接程序。</span><span class="sxs-lookup"><span data-stu-id="328ae-130">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="328ae-131">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="328ae-131">Office Common API requirement sets</span></span>

<span data-ttu-id="328ae-132">大多数 PowerPoint 加载项功能都来自通用 API 集。</span><span class="sxs-lookup"><span data-stu-id="328ae-132">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="328ae-133">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="328ae-133">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="328ae-134">另请参阅</span><span class="sxs-lookup"><span data-stu-id="328ae-134">See also</span></span>

- [<span data-ttu-id="328ae-135">PowerPoint JavaScript API 参考文档</span><span class="sxs-lookup"><span data-stu-id="328ae-135">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="328ae-136">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="328ae-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="328ae-137">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="328ae-137">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="328ae-138">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="328ae-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)

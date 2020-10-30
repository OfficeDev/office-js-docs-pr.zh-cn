---
title: PowerPoint JavaScript API 要求集
description: 了解有关 PowerPoint JavaScript API 要求集的详细信息。
ms.date: 10/26/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: cf9ab510e4b35a140c77ee958279cb85a2189fa2
ms.sourcegitcommit: a4e09546fd59579439025aca9cc58474b5ae7676
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/27/2020
ms.locfileid: "48774725"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="2bf91-103">PowerPoint JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="2bf91-103">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="2bf91-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="2bf91-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="2bf91-107">下表列出了 PowerPoint 要求集、支持这些要求集的 Office 客户端应用程序，以及这些应用程序的内部版本或发布日期。</span><span class="sxs-lookup"><span data-stu-id="2bf91-107">The following table lists the PowerPoint requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="2bf91-108">要求集</span><span class="sxs-lookup"><span data-stu-id="2bf91-108">Requirement set</span></span>  |  <span data-ttu-id="2bf91-109">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="2bf91-109">Office on Windows</span></span><br><span data-ttu-id="2bf91-110">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2bf91-110">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="2bf91-111">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="2bf91-111">Office on iPad</span></span><br><span data-ttu-id="2bf91-112">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2bf91-112">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="2bf91-113">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="2bf91-113">Office on Mac</span></span><br><span data-ttu-id="2bf91-114">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="2bf91-114">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="2bf91-115">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="2bf91-115">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| [<span data-ttu-id="2bf91-116">预览</span><span class="sxs-lookup"><span data-stu-id="2bf91-116">Preview</span></span>](powerpoint-preview-apis.md)  | <span data-ttu-id="2bf91-117">请使用最新的 Office 版本来试用预览 API（你可能需要加入 [Office 预览体验成员计划](https://insider.office.com)）。</span><span class="sxs-lookup"><span data-stu-id="2bf91-117">Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://insider.office.com)).</span></span> |
| <span data-ttu-id="2bf91-118">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="2bf91-118">PowerPointApi 1.1</span></span> | <span data-ttu-id="2bf91-119">版本 1810（内部版本 11001.20074）或更高版本</span><span class="sxs-lookup"><span data-stu-id="2bf91-119">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="2bf91-120">2.17 或更高版本</span><span class="sxs-lookup"><span data-stu-id="2bf91-120">2.17 or later</span></span> | <span data-ttu-id="2bf91-121">16.19 或更高版本</span><span class="sxs-lookup"><span data-stu-id="2bf91-121">16.19 or later</span></span> | <span data-ttu-id="2bf91-122">2018 年 10 月</span><span class="sxs-lookup"><span data-stu-id="2bf91-122">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="2bf91-123">Office 版本和内部版本号</span><span class="sxs-lookup"><span data-stu-id="2bf91-123">Office versions and build numbers</span></span>

<span data-ttu-id="2bf91-124">有关 Office 版本和内部版本号的详细信息，请参阅：</span><span class="sxs-lookup"><span data-stu-id="2bf91-124">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="2bf91-125">PowerPoint JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="2bf91-125">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="2bf91-126">PowerPoint JavaScript API 1.1 包含[用于创建新演示文稿的单一 API](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-)。</span><span class="sxs-lookup"><span data-stu-id="2bf91-126">PowerPoint JavaScript API 1.1 contains a [single API to create a new presentation](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-).</span></span> <span data-ttu-id="2bf91-127">有关 API 的详细信息，请参阅[创建演示文稿](../../powerpoint/powerpoint-add-ins.md#create-a-presentation)。</span><span class="sxs-lookup"><span data-stu-id="2bf91-127">For details about the API, see [Create a presentation](../../powerpoint/powerpoint-add-ins.md#create-a-presentation).</span></span>

## <a name="how-to-use-powerpoint-requirement-sets-at-runtime-and-in-the-manifest"></a><span data-ttu-id="2bf91-128">如何在运行时和清单中使用 PowerPoint 要求集</span><span class="sxs-lookup"><span data-stu-id="2bf91-128">How to use PowerPoint requirement sets at runtime and in the manifest</span></span>

> [!NOTE]
> <span data-ttu-id="2bf91-129">本节假定你熟悉 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)和[指定 Office 应用程序和 API 要求集](../../develop/specify-office-hosts-and-api-requirements.md)处的要求集概述。</span><span class="sxs-lookup"><span data-stu-id="2bf91-129">This section assumes you're familiar with the overview of requirement sets at [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md) and [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md).</span></span>

<span data-ttu-id="2bf91-130">要求集是指各组已命名的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="2bf91-130">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="2bf91-131">Office 加载项可以执行运行时检查或使用清单中指定的要求集确定 Office 应用程序是否支持加载项所需的 API。</span><span class="sxs-lookup"><span data-stu-id="2bf91-131">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office application supports the APIs that the add-in needs.</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="2bf91-132">在运行时检查要求集支持</span><span class="sxs-lookup"><span data-stu-id="2bf91-132">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="2bf91-133">以下代码示例显示如何确定运行加载项的 Office 应用程序是否支持指定的 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="2bf91-133">The following code sample shows how to determine whether the Office application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
} else {
  // Provide alternate flow/logic.
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="2bf91-134">在清单中定义要求集支持</span><span class="sxs-lookup"><span data-stu-id="2bf91-134">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="2bf91-135">可以在加载项清单中使用[要求元素](../manifest/requirements.md)指定加载项要求激活的最小要求集和/或 API 方法。</span><span class="sxs-lookup"><span data-stu-id="2bf91-135">You can use the [Requirements element](../manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="2bf91-136">如果 Office 应用程序或平台不支持清单的 `Requirements` 元素中指定的要求集或 API 方法，则加载项将不会在该应用程序或平台中运行，也不会出现在“ **我的加载项** ”中显示的加载项列表中。如果你的加载项需要特定要求集以实现完整功能，但是即使在不支持该要求集的平台上也可以为用户提供值，则建议在运行时按照上述方式检查要求支持，而不是在清单中定义要求集支持。</span><span class="sxs-lookup"><span data-stu-id="2bf91-136">If the Office application or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that application or platform, and it won't display in the list of add-ins that are shown in **My Add-ins** . If your add-in requires a specific requirement set for full functionality, but it can provide value even to users on platforms that do not support the requirement set, we recommend that you check for requirement support at runtime as described above, instead of defining requirement set support in the manifest.</span></span>

<span data-ttu-id="2bf91-137">以下代码示例显示加载项清单中的 `Requirements` 元素，该元素指定应在支持 PowerPointApi 要求集版本 1.1 或更高版本的所有 Office 客户端应用程序中加载该加载项。</span><span class="sxs-lookup"><span data-stu-id="2bf91-137">The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office client applications that support PowerPointApi requirement set version 1.1 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="2bf91-138">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="2bf91-138">Office Common API requirement sets</span></span>

<span data-ttu-id="2bf91-139">大多数 PowerPoint 加载项功能都来自通用 API 集。</span><span class="sxs-lookup"><span data-stu-id="2bf91-139">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="2bf91-140">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="2bf91-140">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="2bf91-141">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2bf91-141">See also</span></span>

- [<span data-ttu-id="2bf91-142">PowerPoint JavaScript API 参考文档</span><span class="sxs-lookup"><span data-stu-id="2bf91-142">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="2bf91-143">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="2bf91-143">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="2bf91-144">指定 Office 应用程序和 API 要求</span><span class="sxs-lookup"><span data-stu-id="2bf91-144">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="2bf91-145">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="2bf91-145">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)

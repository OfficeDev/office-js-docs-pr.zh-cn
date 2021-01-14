---
title: PowerPoint JavaScript API 要求集
description: 了解有关 PowerPoint JavaScript API 要求集的详细信息。
ms.date: 01/08/2021
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: 63f11f1810b38471a27766843f512da193394838
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2021
ms.locfileid: "49840081"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="31465-103">PowerPoint JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="31465-103">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="31465-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="31465-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="31465-107">下表列出了 PowerPoint 要求集、支持这些要求集的 Office 客户端应用程序，以及这些应用程序的内部版本或发布日期。</span><span class="sxs-lookup"><span data-stu-id="31465-107">The following table lists the PowerPoint requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="31465-108">要求集</span><span class="sxs-lookup"><span data-stu-id="31465-108">Requirement set</span></span>  |  <span data-ttu-id="31465-109">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="31465-109">Office on Windows</span></span><br><span data-ttu-id="31465-110">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="31465-110">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="31465-111">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="31465-111">Office on iPad</span></span><br><span data-ttu-id="31465-112">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="31465-112">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="31465-113">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="31465-113">Office on Mac</span></span><br><span data-ttu-id="31465-114">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="31465-114">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="31465-115">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="31465-115">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| [<span data-ttu-id="31465-116">PowerPointApi 1.2</span><span class="sxs-lookup"><span data-stu-id="31465-116">PowerPointApi 1.2</span></span>](powerpoint-api-1-2-requirement-set.md)  | <span data-ttu-id="31465-117">版本 2011（内部版本 13426.20184）或更高版本</span><span class="sxs-lookup"><span data-stu-id="31465-117">Version 2011 (Build 13426.20184) or later</span></span>| <span data-ttu-id="31465-118">尚不可以</span><span class="sxs-lookup"><span data-stu-id="31465-118">not yet</span></span><br><span data-ttu-id="31465-119">支持</span><span class="sxs-lookup"><span data-stu-id="31465-119">supported</span></span> | <span data-ttu-id="31465-120">16.43 或更高版本</span><span class="sxs-lookup"><span data-stu-id="31465-120">16.43 or later</span></span> | <span data-ttu-id="31465-121">2020 年 10 月</span><span class="sxs-lookup"><span data-stu-id="31465-121">October 2020</span></span> |
| [<span data-ttu-id="31465-122">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="31465-122">PowerPointApi 1.1</span></span>](powerpoint-api-1-1-requirement-set.md) | <span data-ttu-id="31465-123">版本 1810（内部版本 11001.20074）或更高版本</span><span class="sxs-lookup"><span data-stu-id="31465-123">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="31465-124">2.17 或更高版本</span><span class="sxs-lookup"><span data-stu-id="31465-124">2.17 or later</span></span> | <span data-ttu-id="31465-125">16.19 或更高版本</span><span class="sxs-lookup"><span data-stu-id="31465-125">16.19 or later</span></span> | <span data-ttu-id="31465-126">2018 年 10 月</span><span class="sxs-lookup"><span data-stu-id="31465-126">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="31465-127">Office 版本和内部版本号</span><span class="sxs-lookup"><span data-stu-id="31465-127">Office versions and build numbers</span></span>

<span data-ttu-id="31465-128">有关 Office 版本和内部版本号的详细信息，请参阅：</span><span class="sxs-lookup"><span data-stu-id="31465-128">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="31465-129">PowerPoint JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="31465-129">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="31465-130">PowerPoint JavaScript API 1.1 包含[用于创建新演示文稿的单一 API](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-)。</span><span class="sxs-lookup"><span data-stu-id="31465-130">PowerPoint JavaScript API 1.1 contains a [single API to create a new presentation](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-).</span></span> <span data-ttu-id="31465-131">有关 API 的详细信息，请参阅[创建演示文稿](../../powerpoint/powerpoint-add-ins.md#create-a-presentation)。</span><span class="sxs-lookup"><span data-stu-id="31465-131">For details about the API, see [Create a presentation](../../powerpoint/powerpoint-add-ins.md#create-a-presentation).</span></span>

## <a name="powerpoint-javascript-api-12"></a><span data-ttu-id="31465-132">PowerPoint JavaScript API 1.2</span><span class="sxs-lookup"><span data-stu-id="31465-132">PowerPoint JavaScript API 1.2</span></span>

<span data-ttu-id="31465-133">PowerPoint JavaScript API 1.2 增加了对将其他 PowerPoint 演示文稿中的幻灯片插入当前演示文稿以及删除幻灯片的支持.</span><span class="sxs-lookup"><span data-stu-id="31465-133">PowerPoint JavaScript API 1.2 adds support for inserting slides from another PowerPoint presentation into the current presentation and for deleting slides.</span></span> <span data-ttu-id="31465-134">有关 API 的详细信息，请参阅[在 PowerPoint 演示文稿中插入和删除幻灯片](../../powerpoint/insert-slides-into-presentation.md)。</span><span class="sxs-lookup"><span data-stu-id="31465-134">For details about the APIs, see [Insert and delete slides in a PowerPoint presentation](../../powerpoint/insert-slides-into-presentation.md).</span></span>

## <a name="how-to-use-powerpoint-requirement-sets-at-runtime-and-in-the-manifest"></a><span data-ttu-id="31465-135">如何在运行时和清单中使用 PowerPoint 要求集</span><span class="sxs-lookup"><span data-stu-id="31465-135">How to use PowerPoint requirement sets at runtime and in the manifest</span></span>

> [!NOTE]
> <span data-ttu-id="31465-136">本节假定你熟悉 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)和[指定 Office 应用程序和 API 要求集](../../develop/specify-office-hosts-and-api-requirements.md)处的要求集概述。</span><span class="sxs-lookup"><span data-stu-id="31465-136">This section assumes you're familiar with the overview of requirement sets at [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md) and [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md).</span></span>

<span data-ttu-id="31465-137">要求集是指各组已命名的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="31465-137">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="31465-138">Office 加载项可以执行运行时检查或使用清单中指定的要求集确定 Office 应用程序是否支持加载项所需的 API。</span><span class="sxs-lookup"><span data-stu-id="31465-138">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office application supports the APIs that the add-in needs.</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="31465-139">在运行时检查要求集支持</span><span class="sxs-lookup"><span data-stu-id="31465-139">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="31465-140">以下代码示例显示如何确定运行加载项的 Office 应用程序是否支持指定的 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="31465-140">The following code sample shows how to determine whether the Office application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
} else {
  // Provide alternate flow/logic.
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="31465-141">在清单中定义要求集支持</span><span class="sxs-lookup"><span data-stu-id="31465-141">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="31465-142">可以在加载项清单中使用[要求元素](../manifest/requirements.md)指定加载项要求激活的最小要求集和/或 API 方法。</span><span class="sxs-lookup"><span data-stu-id="31465-142">You can use the [Requirements element](../manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="31465-143">如果 Office 应用程序或平台不支持清单的 `Requirements` 元素中指定的要求集或 API 方法，则加载项将不会在该应用程序或平台中运行，也不会出现在“**我的加载项**”中显示的加载项列表中。如果你的加载项需要特定要求集以实现完整功能，但是即使在不支持该要求集的平台上也可以为用户提供值，则建议在运行时按照上述方式检查要求支持，而不是在清单中定义要求集支持。</span><span class="sxs-lookup"><span data-stu-id="31465-143">If the Office application or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that application or platform, and it won't display in the list of add-ins that are shown in **My Add-ins**. If your add-in requires a specific requirement set for full functionality, but it can provide value even to users on platforms that don't support the requirement set, we recommend that you check for requirement support at runtime as described above, instead of defining requirement set support in the manifest.</span></span>

<span data-ttu-id="31465-144">以下代码示例显示加载项清单中的 `Requirements` 元素，该元素指定应在支持 PowerPointApi 要求集版本 1.1 或更高版本的所有 Office 客户端应用程序中加载该加载项。</span><span class="sxs-lookup"><span data-stu-id="31465-144">The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office client applications that support PowerPointApi requirement set version 1.1 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="31465-145">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="31465-145">Office Common API requirement sets</span></span>

<span data-ttu-id="31465-146">大多数 PowerPoint 加载项功能都来自通用 API 集。</span><span class="sxs-lookup"><span data-stu-id="31465-146">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="31465-147">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="31465-147">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="31465-148">另请参阅</span><span class="sxs-lookup"><span data-stu-id="31465-148">See also</span></span>

- [<span data-ttu-id="31465-149">PowerPoint JavaScript API 参考文档</span><span class="sxs-lookup"><span data-stu-id="31465-149">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="31465-150">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="31465-150">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="31465-151">指定 Office 应用程序和 API 要求</span><span class="sxs-lookup"><span data-stu-id="31465-151">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="31465-152">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="31465-152">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)

---
title: OneNote JavaScript API 要求集
description: 了解有关 OneNote JavaScript API 要求集的详细信息。
ms.date: 08/24/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: ecdb26edca54758540688ba03b1d9c1eec14e739
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350188"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="b4f31-103">OneNote JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="b4f31-103">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="b4f31-p101">要求集是指已命名的 API 成员组。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="b4f31-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="b4f31-107">下表列出了 OneNote 要求集、支持这些要求集的 Office 客户端应用程序，以及这些应用程序的内部版本或发布日期。</span><span class="sxs-lookup"><span data-stu-id="b4f31-107">The following table lists the OneNote requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="b4f31-108">要求集</span><span class="sxs-lookup"><span data-stu-id="b4f31-108">Requirement set</span></span>  |  <span data-ttu-id="b4f31-109">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="b4f31-109">Office on the web</span></span> |
|:-----|:-----|
| [<span data-ttu-id="b4f31-110">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="b4f31-110">OneNoteApi 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1&preserve-view=true)  | <span data-ttu-id="b4f31-111">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="b4f31-111">September 2016</span></span> |  

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="b4f31-112">OneNote JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="b4f31-112">OneNote JavaScript API 1.1</span></span>

<span data-ttu-id="b4f31-p102">OneNote JavaScript API 1.1 是首版 API。有关此 API 的详细信息，请参阅 [OneNote JavaScript API 编程概述](../../onenote/onenote-add-ins-programming-overview.md)。</span><span class="sxs-lookup"><span data-stu-id="b4f31-p102">OneNote JavaScript API 1.1 is the first version of the API. For details about the API, see the [OneNote JavaScript API programming overview](../../onenote/onenote-add-ins-programming-overview.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="b4f31-115">运行时要求支持检查</span><span class="sxs-lookup"><span data-stu-id="b4f31-115">Runtime requirement support check</span></span>

<span data-ttu-id="b4f31-116">在运行时，加载项可以执行下列检查，确定特定 Office 应用程序是否支持 API 要求集：</span><span class="sxs-lookup"><span data-stu-id="b4f31-116">At runtime, add-ins can check if a particular Office application supports an API requirement set by doing the following:</span></span>

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="b4f31-117">基于清单的要求支持检查</span><span class="sxs-lookup"><span data-stu-id="b4f31-117">Manifest-based requirement support check</span></span>

<span data-ttu-id="b4f31-p103">只能使用外接程序清单中的 `Requirements` 元素指定外接程序必须使用的关键要求集或 API 成员。如果 Office 应用程序或平台不支持在 `Requirements` 元素中指定的要求集或 API 成员，则外接程序将无法在该应用程序或平台上运行，并且不会显示在“我的外接程序”中。</span><span class="sxs-lookup"><span data-stu-id="b4f31-p103">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office application or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that application or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="b4f31-120">下面的代码示例展示了加载所有支持第 1.1 版 OneNoteApi 要求集的 Office 客户端应用程序的加载项。</span><span class="sxs-lookup"><span data-stu-id="b4f31-120">The following code example shows an add-in that loads in all Office client applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="b4f31-121">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="b4f31-121">Office Common API requirement sets</span></span>

<span data-ttu-id="b4f31-122">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="b4f31-122">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="b4f31-123">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b4f31-123">See also</span></span>

- [<span data-ttu-id="b4f31-124">OneNote JavaScript API 参考文档</span><span class="sxs-lookup"><span data-stu-id="b4f31-124">OneNote JavaScript API reference documentation</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="b4f31-125">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="b4f31-125">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="b4f31-126">指定 Office 应用程序和 API 要求集</span><span class="sxs-lookup"><span data-stu-id="b4f31-126">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="b4f31-127">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="b4f31-127">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)

---
title: OneNote JavaScript API 要求集
description: 了解有关 OneNote JavaScript API 要求集的详细信息
ms.date: 07/17/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 7717ff20fe4a7f29621a30df7d01d122111021db
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717444"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="dd918-103">OneNote JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="dd918-103">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="dd918-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="dd918-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="dd918-107">下表列出了 OneNote 要求集、支持这些要求集的 Office 主机应用程序，以及这些应用程序的内部版本或发布日期。</span><span class="sxs-lookup"><span data-stu-id="dd918-107">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="dd918-108">要求集</span><span class="sxs-lookup"><span data-stu-id="dd918-108">Requirement set</span></span>  |  <span data-ttu-id="dd918-109">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="dd918-109">Office on the web</span></span> |
|:-----|:-----|
| [<span data-ttu-id="dd918-110">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="dd918-110">OneNoteApi 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)  | <span data-ttu-id="dd918-111">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="dd918-111">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="dd918-112">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="dd918-112">Office Common API requirement sets</span></span>

<span data-ttu-id="dd918-113">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="dd918-113">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="dd918-114">OneNote JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="dd918-114">OneNote JavaScript API 1.1</span></span>

<span data-ttu-id="dd918-115">OneNote JavaScript API 1.1 是该 API 的第一版。</span><span class="sxs-lookup"><span data-stu-id="dd918-115">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="dd918-116">有关此 API 的详细信息，请参阅 [OneNote JavaScript API 编程概述](../../onenote/onenote-add-ins-programming-overview.md)。</span><span class="sxs-lookup"><span data-stu-id="dd918-116">For details about the API, see the [OneNote JavaScript API programming overview](../../onenote/onenote-add-ins-programming-overview.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="dd918-117">运行时要求支持检查</span><span class="sxs-lookup"><span data-stu-id="dd918-117">Runtime requirement support check</span></span>

<span data-ttu-id="dd918-118">在运行时，加载项可以执行下列检查，确定特定主机是否支持 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="dd918-118">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="dd918-119">基于清单的要求支持检查</span><span class="sxs-lookup"><span data-stu-id="dd918-119">Manifest-based requirement support check</span></span>

<span data-ttu-id="dd918-120">使用加载项清单中的 `Requirements` 元素指定加载项必须使用的关键要求集或 API 成员。</span><span class="sxs-lookup"><span data-stu-id="dd918-120">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="dd918-121">如果 Office 主机或平台不支持 `Requirements` 元素中指定的要求集或 API 成员，则加载项将无法在该主机或平台上运行，并且不会显示在“我的加载项”中。</span><span class="sxs-lookup"><span data-stu-id="dd918-121">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="dd918-122">下面的代码示例展示了加载所有 支持第 1.1 版 OneNoteApi 要求集的 Office 主机应用程序的外接程序。</span><span class="sxs-lookup"><span data-stu-id="dd918-122">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="dd918-123">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="dd918-123">Office Common API requirement sets</span></span>

<span data-ttu-id="dd918-124">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="dd918-124">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="dd918-125">另请参阅</span><span class="sxs-lookup"><span data-stu-id="dd918-125">See also</span></span>

- [<span data-ttu-id="dd918-126">OneNote JavaScript API 参考文档</span><span class="sxs-lookup"><span data-stu-id="dd918-126">OneNote JavaScript API reference documentation</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="dd918-127">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="dd918-127">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="dd918-128">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="dd918-128">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="dd918-129">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="dd918-129">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)

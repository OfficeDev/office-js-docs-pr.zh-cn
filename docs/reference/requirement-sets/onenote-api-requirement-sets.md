---
title: OneNote JavaScript API 要求集
description: ''
ms.date: 03/19/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: 287e405955477a98854b1df4a81fe90ec16e5bbc
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450098"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="f9643-102">OneNote JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="f9643-102">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="f9643-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="f9643-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="f9643-106">下表列出了 OneNote 要求集、支持这些要求集的 Office 主机应用程序，以及这些应用程序的内部版本或发布日期。</span><span class="sxs-lookup"><span data-stu-id="f9643-106">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="f9643-107">要求集</span><span class="sxs-lookup"><span data-stu-id="f9643-107">Requirement set</span></span>  |  <span data-ttu-id="f9643-108">Office Online</span><span class="sxs-lookup"><span data-stu-id="f9643-108">Office Online</span></span> | 
|:-----|:-----|
| <span data-ttu-id="f9643-109">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f9643-109">OneNoteApi 1.1</span></span>  | <span data-ttu-id="f9643-110">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="f9643-110">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="f9643-111">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="f9643-111">Office Common API requirement sets</span></span>

<span data-ttu-id="f9643-112">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="f9643-112">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="f9643-113">OneNote JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="f9643-113">OneNote JavaScript API 1.1</span></span> 

<span data-ttu-id="f9643-114">OneNote JavaScript API 1.1 是该 API 的第一版。</span><span class="sxs-lookup"><span data-stu-id="f9643-114">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="f9643-115">有关此 API 的详细信息，请参阅 [OneNote JavaScript API 编程概述](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)。</span><span class="sxs-lookup"><span data-stu-id="f9643-115">For details about the API, see the [OneNote JavaScript API programming overview](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="f9643-116">运行时要求支持检查</span><span class="sxs-lookup"><span data-stu-id="f9643-116">Runtime requirement support check</span></span>

<span data-ttu-id="f9643-117">在运行时期间，加载项可以执行下列检查，确定特定主机是否支持 API 要求集：</span><span class="sxs-lookup"><span data-stu-id="f9643-117">During the runtime, add-ins can check if a particular host supports an API requirement set by doing the following-check:</span></span> 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="f9643-118">基于清单的要求支持检查</span><span class="sxs-lookup"><span data-stu-id="f9643-118">Manifest-based requirement support check</span></span>

<span data-ttu-id="f9643-p103">只能使用外接程序清单中的 Requirements 元素指定外接程序必须使用的关键要求集或 API 成员。如果 Office 主机或平台不支持在 Requirements 元素中指定的要求集或 API 成员，则外接程序将无法在该主机或平台上运行，并且不会显示在“我的外接程序”中。</span><span class="sxs-lookup"><span data-stu-id="f9643-p103">Use the Requirements element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the Requirements element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="f9643-121">下面的代码示例展示了加载所有 支持第 1.1 版 OneNoteApi 要求集的 Office 主机应用程序的外接程序。</span><span class="sxs-lookup"><span data-stu-id="f9643-121">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a><span data-ttu-id="f9643-122">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f9643-122">See also</span></span>

- [<span data-ttu-id="f9643-123">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="f9643-123">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="f9643-124">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="f9643-124">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="f9643-125">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="f9643-125">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)

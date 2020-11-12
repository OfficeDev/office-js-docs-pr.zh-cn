---
title: Excel JavaScript API 仅联机要求集
description: 有关 ExcelApiOnline 要求集的详细信息。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 16c96f413424d5fc85a21419fb72cf6580c1ac18
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996527"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a><span data-ttu-id="42701-103">Excel JavaScript API 仅联机要求集</span><span class="sxs-lookup"><span data-stu-id="42701-103">Excel JavaScript API online-only requirement set</span></span>

<span data-ttu-id="42701-104">`ExcelApiOnline`要求集是一个特殊要求集，其中包含仅适用于 web 上的 Excel 的功能。</span><span class="sxs-lookup"><span data-stu-id="42701-104">The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web.</span></span> <span data-ttu-id="42701-105">此要求集中的 Api 被视为生产 Api (不受未记录的行为或结构更改) 针对 web 应用程序上的 Excel。</span><span class="sxs-lookup"><span data-stu-id="42701-105">APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web application.</span></span> <span data-ttu-id="42701-106">`ExcelApiOnline` 被认为是其他平台 (Windows、Mac、iOS) 的 "预览" Api，这些平台可能不支持这些平台。</span><span class="sxs-lookup"><span data-stu-id="42701-106">`ExcelApiOnline` are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.</span></span>

<span data-ttu-id="42701-107">当 `ExcelApiOnline` 所有平台都支持要求集中的 api 时，它们将添加到下一个发布的要求集 (`ExcelApi 1.[NEXT]`) 。</span><span class="sxs-lookup"><span data-stu-id="42701-107">When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`).</span></span> <span data-ttu-id="42701-108">一旦新要求是公共的，将从这些 Api 中删除 `ExcelApiOnline` 。</span><span class="sxs-lookup"><span data-stu-id="42701-108">Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`.</span></span> <span data-ttu-id="42701-109">可将此视为将 API 从预览迁移到发布的类似升级过程。</span><span class="sxs-lookup"><span data-stu-id="42701-109">Think of this as a similar promotion process as an API moving from preview to release.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="42701-110">`ExcelApiOnline` 是最新编号的要求集的超集。</span><span class="sxs-lookup"><span data-stu-id="42701-110">`ExcelApiOnline` is superset of the latest numbered requirement set.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="42701-111">`ExcelApiOnline 1.1` 是仅联机 Api 的唯一版本。</span><span class="sxs-lookup"><span data-stu-id="42701-111">`ExcelApiOnline 1.1` is the only version of the online-only APIs.</span></span> <span data-ttu-id="42701-112">这是因为 web 上的 Excel 将始终有一个版本可供最新版本的用户使用。</span><span class="sxs-lookup"><span data-stu-id="42701-112">This is because Excel on the web will always have a single version available to users that is the latest version.</span></span>

## <a name="recommended-usage"></a><span data-ttu-id="42701-113">建议使用</span><span class="sxs-lookup"><span data-stu-id="42701-113">Recommended usage</span></span>

<span data-ttu-id="42701-114">由于 `ExcelApiOnline` web 上的 Excel 仅支持 api，因此，您的外接程序应检查是否支持要求集，然后再调用这些 api。</span><span class="sxs-lookup"><span data-stu-id="42701-114">Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs.</span></span> <span data-ttu-id="42701-115">这样可以避免在不同的平台上调用仅联机 API。</span><span class="sxs-lookup"><span data-stu-id="42701-115">This avoids calling an online-only API on a different platform.</span></span>

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

<span data-ttu-id="42701-116">一旦 API 位于跨平台要求集，就应删除或编辑该 `isSetSupported` 检查。</span><span class="sxs-lookup"><span data-stu-id="42701-116">Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check.</span></span> <span data-ttu-id="42701-117">这将在其他平台上启用外接程序的功能。</span><span class="sxs-lookup"><span data-stu-id="42701-117">This will enable your add-in's feature on other platforms.</span></span> <span data-ttu-id="42701-118">进行此更改时，请务必在这些平台上测试功能。</span><span class="sxs-lookup"><span data-stu-id="42701-118">Be sure to test the feature on those platforms when making this change.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="42701-119">清单不能指定 `ExcelApiOnline 1.1` 为激活要求。</span><span class="sxs-lookup"><span data-stu-id="42701-119">Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement.</span></span> <span data-ttu-id="42701-120">不是在 [Set 元素](../manifest/set.md)中使用的有效值。</span><span class="sxs-lookup"><span data-stu-id="42701-120">It is not a valid value to use in the [Set element](../manifest/set.md).</span></span>

## <a name="api-list"></a><span data-ttu-id="42701-121">API 列表</span><span class="sxs-lookup"><span data-stu-id="42701-121">API list</span></span>

| <span data-ttu-id="42701-122">Class</span><span class="sxs-lookup"><span data-stu-id="42701-122">Class</span></span> | <span data-ttu-id="42701-123">域</span><span class="sxs-lookup"><span data-stu-id="42701-123">Fields</span></span> | <span data-ttu-id="42701-124">说明</span><span class="sxs-lookup"><span data-stu-id="42701-124">Description</span></span> |
|:---|:---|:---|
|[<span data-ttu-id="42701-125">Range</span><span class="sxs-lookup"><span data-stu-id="42701-125">Range</span></span>](/javascript/api/excel/excel.range)|[<span data-ttu-id="42701-126">getMergedAreas ( # B1 </span><span class="sxs-lookup"><span data-stu-id="42701-126">getMergedAreas()</span></span>](/javascript/api/excel/excel.range#getmergedareas--)|<span data-ttu-id="42701-127">返回一个 RangeAreas 对象，该对象代表此区域中的合并区域。</span><span class="sxs-lookup"><span data-stu-id="42701-127">Returns a RangeAreas object that represents the merged areas in this range.</span></span>|

## <a name="see-also"></a><span data-ttu-id="42701-128">另请参阅</span><span class="sxs-lookup"><span data-stu-id="42701-128">See also</span></span>

- [<span data-ttu-id="42701-129">Excel JavaScript API 参考文档</span><span class="sxs-lookup"><span data-stu-id="42701-129">Excel JavaScript API Reference Documentation</span></span>](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [<span data-ttu-id="42701-130">Excel JavaScript 预览 API</span><span class="sxs-lookup"><span data-stu-id="42701-130">Excel JavaScript preview APIs</span></span>](excel-preview-apis.md)
- [<span data-ttu-id="42701-131">Excel JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="42701-131">Excel JavaScript API requirement sets</span></span>](excel-api-requirement-sets.md)

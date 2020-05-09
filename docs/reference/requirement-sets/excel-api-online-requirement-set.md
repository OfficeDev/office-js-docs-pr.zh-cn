---
title: Excel JavaScript API 仅联机要求集
description: 有关 ExcelApiOnline 要求集的详细信息
ms.date: 05/06/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f177e0107de7172c350f94c3a022cb3e0db5c6f5
ms.sourcegitcommit: 735bf94ac3c838f580a992e7ef074dbc8be2b0ea
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/08/2020
ms.locfileid: "44170784"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a><span data-ttu-id="69f20-103">Excel JavaScript API 仅联机要求集</span><span class="sxs-lookup"><span data-stu-id="69f20-103">Excel JavaScript API online-only requirement set</span></span>

<span data-ttu-id="69f20-104">`ExcelApiOnline`要求集是一个特殊要求集，其中包含仅适用于 web 上的 Excel 的功能。</span><span class="sxs-lookup"><span data-stu-id="69f20-104">The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web.</span></span> <span data-ttu-id="69f20-105">此要求集中的 Api 被认为是针对 web 主机上的 Excel 的生产 Api （不受未记录的行为或结构更改）。</span><span class="sxs-lookup"><span data-stu-id="69f20-105">APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web host.</span></span> <span data-ttu-id="69f20-106">`ExcelApiOnline`被视为针对其他平台（Windows、Mac、iOS）的 "预览" Api，这些平台可能不支持这些平台。</span><span class="sxs-lookup"><span data-stu-id="69f20-106">`ExcelApiOnline` are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.</span></span>

<span data-ttu-id="69f20-107">当在所有平台`ExcelApiOnline`上支持要求集中的 api 时，它们将添加到下一个发布的要求集`ExcelApi 1.[NEXT]`（）。</span><span class="sxs-lookup"><span data-stu-id="69f20-107">When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`).</span></span> <span data-ttu-id="69f20-108">一旦新要求是公共的，将从这些 Api 中`ExcelApiOnline`删除。</span><span class="sxs-lookup"><span data-stu-id="69f20-108">Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`.</span></span> <span data-ttu-id="69f20-109">可将此视为将 API 从预览迁移到发布的类似升级过程。</span><span class="sxs-lookup"><span data-stu-id="69f20-109">Think of this as a similar promotion process as an API moving from preview to release.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="69f20-110">`ExcelApiOnline`是最新编号的要求集的超集。</span><span class="sxs-lookup"><span data-stu-id="69f20-110">`ExcelApiOnline` is superset of the latest numbered requirement set.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="69f20-111">`ExcelApiOnline 1.1`是仅联机 Api 的唯一版本。</span><span class="sxs-lookup"><span data-stu-id="69f20-111">`ExcelApiOnline 1.1` is the only version of the online-only APIs.</span></span> <span data-ttu-id="69f20-112">这是因为 web 上的 Excel 将始终有一个版本可供最新版本的用户使用。</span><span class="sxs-lookup"><span data-stu-id="69f20-112">This is because Excel on the web will always have a single version available to users that is the latest version.</span></span>

## <a name="recommended-usage"></a><span data-ttu-id="69f20-113">建议使用</span><span class="sxs-lookup"><span data-stu-id="69f20-113">Recommended usage</span></span>

<span data-ttu-id="69f20-114">由于`ExcelApiOnline` web 上的 Excel 仅支持 api，因此，您的外接程序应检查是否支持要求集，然后再调用这些 api。</span><span class="sxs-lookup"><span data-stu-id="69f20-114">Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs.</span></span> <span data-ttu-id="69f20-115">这样可以避免在不同的平台上调用仅联机 API。</span><span class="sxs-lookup"><span data-stu-id="69f20-115">This avoids calling an online-only API on a different platform.</span></span>

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

<span data-ttu-id="69f20-116">一旦 API 位于跨平台要求集，就应删除或编辑该`isSetSupported`检查。</span><span class="sxs-lookup"><span data-stu-id="69f20-116">Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check.</span></span> <span data-ttu-id="69f20-117">这将在其他平台上启用外接程序的功能。</span><span class="sxs-lookup"><span data-stu-id="69f20-117">This will enable your add-in's feature on other platforms.</span></span> <span data-ttu-id="69f20-118">进行此更改时，请务必在这些平台上测试功能。</span><span class="sxs-lookup"><span data-stu-id="69f20-118">Be sure to test the feature on those platforms when making this change.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="69f20-119">清单不能指定`ExcelApiOnline 1.1`为激活要求。</span><span class="sxs-lookup"><span data-stu-id="69f20-119">Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement.</span></span> <span data-ttu-id="69f20-120">不是在[Set 元素](../manifest/set.md)中使用的有效值。</span><span class="sxs-lookup"><span data-stu-id="69f20-120">It is not a valid value to use in the [Set element](../manifest/set.md).</span></span>

## <a name="api-list"></a><span data-ttu-id="69f20-121">API 列表</span><span class="sxs-lookup"><span data-stu-id="69f20-121">API list</span></span>

<span data-ttu-id="69f20-122">下面的 Api 当前可用于 web 上的 Excel，作为`ExcelApiOnline 1.1`要求集的一部分。</span><span class="sxs-lookup"><span data-stu-id="69f20-122">The following APIs are currently available for Excel on the web as part of the `ExcelApiOnline 1.1` requirement set.</span></span>

| <span data-ttu-id="69f20-123">Class</span><span class="sxs-lookup"><span data-stu-id="69f20-123">Class</span></span> | <span data-ttu-id="69f20-124">域</span><span class="sxs-lookup"><span data-stu-id="69f20-124">Fields</span></span> | <span data-ttu-id="69f20-125">说明</span><span class="sxs-lookup"><span data-stu-id="69f20-125">Description</span></span> |
|:---|:---|:---|
|[<span data-ttu-id="69f20-126">ChartAxisTitle</span><span class="sxs-lookup"><span data-stu-id="69f20-126">ChartAxisTitle</span></span>](/javascript/api/excel/excel.chartaxistitle)|[<span data-ttu-id="69f20-127">textOrientation</span><span class="sxs-lookup"><span data-stu-id="69f20-127">textOrientation</span></span>](/javascript/api/excel/excel.chartaxistitle#textorientation)|<span data-ttu-id="69f20-128">指定文本面向图表轴标题的角度。</span><span class="sxs-lookup"><span data-stu-id="69f20-128">Specifies the angle to which the text is oriented for the chart axis title.</span></span> <span data-ttu-id="69f20-129">该值应为-90 到90的整数或垂直方向的文本的整数180。</span><span class="sxs-lookup"><span data-stu-id="69f20-129">The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.</span></span>|
|[<span data-ttu-id="69f20-130">PivotTableScopedCollection</span><span class="sxs-lookup"><span data-stu-id="69f20-130">PivotTableScopedCollection</span></span>](/javascript/api/excel/excel.pivottablescopedcollection)|[<span data-ttu-id="69f20-131">getCount()</span><span class="sxs-lookup"><span data-stu-id="69f20-131">getCount()</span></span>](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|<span data-ttu-id="69f20-132">获取集合中的数据透视表的数目。</span><span class="sxs-lookup"><span data-stu-id="69f20-132">Gets the number of PivotTables in the collection.</span></span>|
||[<span data-ttu-id="69f20-133">getFirst()</span><span class="sxs-lookup"><span data-stu-id="69f20-133">getFirst()</span></span>](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|<span data-ttu-id="69f20-134">获取集合中的第一个数据透视表。</span><span class="sxs-lookup"><span data-stu-id="69f20-134">Gets the first PivotTable in the collection.</span></span> <span data-ttu-id="69f20-135">集合中的数据透视表按从上到下、从左到右的顺序排序，因此左上角的表格是集合中的第一个数据透视表。</span><span class="sxs-lookup"><span data-stu-id="69f20-135">The PivotTables in the collection are sorted top to bottom and left to right, such that top-left table is the first PivotTable in the collection.</span></span>|
||[<span data-ttu-id="69f20-136">getItem(key: string)</span><span class="sxs-lookup"><span data-stu-id="69f20-136">getItem(key: string)</span></span>](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|<span data-ttu-id="69f20-137">按名称获取 PivotTable 对象。</span><span class="sxs-lookup"><span data-stu-id="69f20-137">Gets a PivotTable by name.</span></span>|
||[<span data-ttu-id="69f20-138">getItemOrNullObject(name: string)</span><span class="sxs-lookup"><span data-stu-id="69f20-138">getItemOrNullObject(name: string)</span></span>](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|<span data-ttu-id="69f20-139">按 PivotTable 对象的名称获取此对象。</span><span class="sxs-lookup"><span data-stu-id="69f20-139">Gets a PivotTable by name.</span></span> <span data-ttu-id="69f20-140">如果没有 PivotTable 对象，将返回 NULL 对象。</span><span class="sxs-lookup"><span data-stu-id="69f20-140">If the PivotTable does not exist, will return a null object.</span></span>|
||[<span data-ttu-id="69f20-141">items</span><span class="sxs-lookup"><span data-stu-id="69f20-141">items</span></span>](/javascript/api/excel/excel.pivottablescopedcollection#items)|<span data-ttu-id="69f20-142">获取此集合中已加载的子项。</span><span class="sxs-lookup"><span data-stu-id="69f20-142">Gets the loaded child items in this collection.</span></span>|
|[<span data-ttu-id="69f20-143">区域</span><span class="sxs-lookup"><span data-stu-id="69f20-143">Range</span></span>](/javascript/api/excel/excel.range)|[<span data-ttu-id="69f20-144">getPivotTables （fullyContained？：布尔值）</span><span class="sxs-lookup"><span data-stu-id="69f20-144">getPivotTables(fullyContained?: boolean)</span></span>](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|<span data-ttu-id="69f20-145">获取与区域重叠的数据透视表的限定集合。</span><span class="sxs-lookup"><span data-stu-id="69f20-145">Gets a scoped collection of PivotTables that overlap with the range.</span></span>|

## <a name="see-also"></a><span data-ttu-id="69f20-146">另请参阅</span><span class="sxs-lookup"><span data-stu-id="69f20-146">See also</span></span>

- [<span data-ttu-id="69f20-147">Excel JavaScript API 参考文档</span><span class="sxs-lookup"><span data-stu-id="69f20-147">Excel JavaScript API Reference Documentation</span></span>](/javascript/api/excel?view=excel-js-online)
- [<span data-ttu-id="69f20-148">Excel JavaScript 预览 API</span><span class="sxs-lookup"><span data-stu-id="69f20-148">Excel JavaScript preview APIs</span></span>](./excel-preview-apis.md)
- [<span data-ttu-id="69f20-149">Excel JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="69f20-149">Excel JavaScript API requirement sets</span></span>](./excel-api-requirement-sets.md)
---
title: ExcelJavaScript API 仅联机要求集
description: 有关 ExcelApiOnline 要求集的详细信息。
ms.date: 07/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ef4831cf6a6f9be1a5413c89ae0f971bef51a9b1
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290801"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a><span data-ttu-id="23c8e-103">ExcelJavaScript API 仅联机要求集</span><span class="sxs-lookup"><span data-stu-id="23c8e-103">Excel JavaScript API online-only requirement set</span></span>

<span data-ttu-id="23c8e-104">要求 `ExcelApiOnline` 集是一个特殊要求集，其中包含仅适用于Excel web 版。</span><span class="sxs-lookup"><span data-stu-id="23c8e-104">The `ExcelApiOnline` requirement set is a special requirement set that includes features that are only available for Excel on the web.</span></span> <span data-ttu-id="23c8e-105">此要求集内 API 被视为生产 API， (应用程序未记录的行为或) 更改Excel web 版 API。</span><span class="sxs-lookup"><span data-stu-id="23c8e-105">APIs in this requirement set are considered to be production APIs (not subject to undocumented behavioral or structural changes) for the Excel on the web application.</span></span> <span data-ttu-id="23c8e-106">`ExcelApiOnline`API 被视为适用于其他平台（如 (Windows、Mac、iOS) ）的"预览"API，可能不受这些平台的任何支持。</span><span class="sxs-lookup"><span data-stu-id="23c8e-106">`ExcelApiOnline` APIs are considered to be "preview" APIs for other platforms (Windows, Mac, iOS) and may not be supported by any of those platforms.</span></span>

<span data-ttu-id="23c8e-107">当所有平台都支持要求集内 API 时，它们将被添加到下一个发布的要求 `ExcelApiOnline` `ExcelApi 1.[NEXT]` () 。</span><span class="sxs-lookup"><span data-stu-id="23c8e-107">When APIs in the `ExcelApiOnline` requirement set are supported across all platforms, they will added to the next released requirement set (`ExcelApi 1.[NEXT]`).</span></span> <span data-ttu-id="23c8e-108">一旦该新要求公开，将从 中删除这些 `ExcelApiOnline` API。</span><span class="sxs-lookup"><span data-stu-id="23c8e-108">Once that new requirement is public, those APIs will be removed from `ExcelApiOnline`.</span></span> <span data-ttu-id="23c8e-109">将此过程视为从预览版移动到发布的 API 的类似推广过程。</span><span class="sxs-lookup"><span data-stu-id="23c8e-109">Think of this as a similar promotion process to an API moving from preview to release.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="23c8e-110">`ExcelApiOnline` 是最新编号要求集的超集。</span><span class="sxs-lookup"><span data-stu-id="23c8e-110">`ExcelApiOnline` is a superset of the latest numbered requirement set.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="23c8e-111">`ExcelApiOnline 1.1` 是仅联机 API 的唯一版本。</span><span class="sxs-lookup"><span data-stu-id="23c8e-111">`ExcelApiOnline 1.1` is the only version of the online-only APIs.</span></span> <span data-ttu-id="23c8e-112">这是因为Excel web 版将始终为最新版本的用户提供单个版本。</span><span class="sxs-lookup"><span data-stu-id="23c8e-112">This is because Excel on the web will always have a single version available to users that is the latest version.</span></span>

<span data-ttu-id="23c8e-113">下表提供了 API 的简要摘要，而后续 API 列表表提供了当前 [API](#api-list) `ExcelApiOnline` 的详细列表。</span><span class="sxs-lookup"><span data-stu-id="23c8e-113">The following table provides a concise summary of the APIs, while the subsequent [API list](#api-list) table gives a detailed list of the current `ExcelApiOnline` APIs.</span></span>

| <span data-ttu-id="23c8e-114">功能区域</span><span class="sxs-lookup"><span data-stu-id="23c8e-114">Feature area</span></span> | <span data-ttu-id="23c8e-115">说明</span><span class="sxs-lookup"><span data-stu-id="23c8e-115">Description</span></span> | <span data-ttu-id="23c8e-116">相关对象</span><span class="sxs-lookup"><span data-stu-id="23c8e-116">Relevant objects</span></span> |
|:--- |:--- |:--- |
| <span data-ttu-id="23c8e-117">命名工作表视图</span><span class="sxs-lookup"><span data-stu-id="23c8e-117">Named sheet views</span></span> | <span data-ttu-id="23c8e-118">以编程方式控制每用户工作表视图。</span><span class="sxs-lookup"><span data-stu-id="23c8e-118">Gives programmatic control of per-user worksheet views.</span></span> | [<span data-ttu-id="23c8e-119">NamedSheetView</span><span class="sxs-lookup"><span data-stu-id="23c8e-119">NamedSheetView</span></span>](/javascript/api/excel/excel.namedsheetview) |

## <a name="recommended-usage"></a><span data-ttu-id="23c8e-120">建议的用法</span><span class="sxs-lookup"><span data-stu-id="23c8e-120">Recommended usage</span></span>

<span data-ttu-id="23c8e-121">由于 API 仅受 Excel web 版，因此加载项应在调用这些 API 之前检查 `ExcelApiOnline` 要求集是否受支持。</span><span class="sxs-lookup"><span data-stu-id="23c8e-121">Because `ExcelApiOnline` APIs are only supported by Excel on the web, your add-in should check if the requirement set is supported before calling these APIs.</span></span> <span data-ttu-id="23c8e-122">这可以避免在不同的平台上调用仅联机 API。</span><span class="sxs-lookup"><span data-stu-id="23c8e-122">This avoids calling an online-only API on a different platform.</span></span>

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

<span data-ttu-id="23c8e-123">API 位于跨平台要求集后，应删除或编辑 `isSetSupported` 检查。</span><span class="sxs-lookup"><span data-stu-id="23c8e-123">Once the API is in a cross-platform requirement set, you should remove or edit the `isSetSupported` check.</span></span> <span data-ttu-id="23c8e-124">这将在其他平台上启用外接程序的功能。</span><span class="sxs-lookup"><span data-stu-id="23c8e-124">This will enable your add-in's feature on other platforms.</span></span> <span data-ttu-id="23c8e-125">进行此更改时，请务必在这些平台上测试该功能。</span><span class="sxs-lookup"><span data-stu-id="23c8e-125">Be sure to test the feature on those platforms when making this change.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="23c8e-126">清单不能指定为 `ExcelApiOnline 1.1` 激活要求。</span><span class="sxs-lookup"><span data-stu-id="23c8e-126">Your manifest cannot specify `ExcelApiOnline 1.1` as an activation requirement.</span></span> <span data-ttu-id="23c8e-127">它在 Set 元素中不是有效的 [值](../manifest/set.md)。</span><span class="sxs-lookup"><span data-stu-id="23c8e-127">It is not a valid value to use in the [Set element](../manifest/set.md).</span></span>

## <a name="api-list"></a><span data-ttu-id="23c8e-128">API 列表</span><span class="sxs-lookup"><span data-stu-id="23c8e-128">API list</span></span>

<span data-ttu-id="23c8e-129">下表列出了要求Excel当前包含的 JavaScript `ExcelApiOnline` API。</span><span class="sxs-lookup"><span data-stu-id="23c8e-129">The following table lists the Excel JavaScript APIs currently included in the `ExcelApiOnline` requirement set.</span></span> <span data-ttu-id="23c8e-130">有关所有 JavaScript API Excel的完整列表 (包括 API 和以前发布的 API `ExcelApiOnline`) ，请参阅所有 Excel [JavaScript API。](/javascript/api/excel?view=excel-js-online&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="23c8e-130">For a complete list of all Excel JavaScript APIs (including `ExcelApiOnline` APIs and previously released APIs), see [all Excel JavaScript APIs](/javascript/api/excel?view=excel-js-online&preserve-view=true).</span></span>

| <span data-ttu-id="23c8e-131">类</span><span class="sxs-lookup"><span data-stu-id="23c8e-131">Class</span></span> | <span data-ttu-id="23c8e-132">域</span><span class="sxs-lookup"><span data-stu-id="23c8e-132">Fields</span></span> | <span data-ttu-id="23c8e-133">说明</span><span class="sxs-lookup"><span data-stu-id="23c8e-133">Description</span></span> |
|:---|:---|:---|
|[<span data-ttu-id="23c8e-134">NamedSheetView</span><span class="sxs-lookup"><span data-stu-id="23c8e-134">NamedSheetView</span></span>](/javascript/api/excel/excel.namedsheetview)|[<span data-ttu-id="23c8e-135">activate()</span><span class="sxs-lookup"><span data-stu-id="23c8e-135">activate()</span></span>](/javascript/api/excel/excel.namedsheetview#activate--)|<span data-ttu-id="23c8e-136">激活此工作表视图。</span><span class="sxs-lookup"><span data-stu-id="23c8e-136">Activates this sheet view.</span></span>|
||[<span data-ttu-id="23c8e-137">delete()</span><span class="sxs-lookup"><span data-stu-id="23c8e-137">delete()</span></span>](/javascript/api/excel/excel.namedsheetview#delete--)|<span data-ttu-id="23c8e-138">从工作表中删除工作表视图。</span><span class="sxs-lookup"><span data-stu-id="23c8e-138">Removes the sheet view from the worksheet.</span></span>|
||[<span data-ttu-id="23c8e-139">duplicate (name？： string) </span><span class="sxs-lookup"><span data-stu-id="23c8e-139">duplicate(name?: string)</span></span>](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|<span data-ttu-id="23c8e-140">创建此工作表视图的副本。</span><span class="sxs-lookup"><span data-stu-id="23c8e-140">Creates a copy of this sheet view.</span></span>|
||[<span data-ttu-id="23c8e-141">name</span><span class="sxs-lookup"><span data-stu-id="23c8e-141">name</span></span>](/javascript/api/excel/excel.namedsheetview#name)|<span data-ttu-id="23c8e-142">获取或设置工作表视图的名称。</span><span class="sxs-lookup"><span data-stu-id="23c8e-142">Gets or sets the name of the sheet view.</span></span>|
|[<span data-ttu-id="23c8e-143">NamedSheetViewCollection</span><span class="sxs-lookup"><span data-stu-id="23c8e-143">NamedSheetViewCollection</span></span>](/javascript/api/excel/excel.namedsheetviewcollection)|[<span data-ttu-id="23c8e-144">add(name: string)</span><span class="sxs-lookup"><span data-stu-id="23c8e-144">add(name: string)</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|<span data-ttu-id="23c8e-145">创建具有给定名称的新工作表视图。</span><span class="sxs-lookup"><span data-stu-id="23c8e-145">Creates a new sheet view with the given name.</span></span>|
||[<span data-ttu-id="23c8e-146">enterTemporary () </span><span class="sxs-lookup"><span data-stu-id="23c8e-146">enterTemporary()</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|<span data-ttu-id="23c8e-147">创建并激活新的临时工作表视图。</span><span class="sxs-lookup"><span data-stu-id="23c8e-147">Creates and activates a new temporary sheet view.</span></span>|
||[<span data-ttu-id="23c8e-148">exit () </span><span class="sxs-lookup"><span data-stu-id="23c8e-148">exit()</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|<span data-ttu-id="23c8e-149">退出当前活动的工作表视图。</span><span class="sxs-lookup"><span data-stu-id="23c8e-149">Exits the currently active sheet view.</span></span>|
||[<span data-ttu-id="23c8e-150">getActive () </span><span class="sxs-lookup"><span data-stu-id="23c8e-150">getActive()</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|<span data-ttu-id="23c8e-151">获取工作表当前的活动工作表视图。</span><span class="sxs-lookup"><span data-stu-id="23c8e-151">Gets the worksheet's currently active sheet view.</span></span>|
||[<span data-ttu-id="23c8e-152">getCount()</span><span class="sxs-lookup"><span data-stu-id="23c8e-152">getCount()</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|<span data-ttu-id="23c8e-153">获取此工作表中的工作表视图数。</span><span class="sxs-lookup"><span data-stu-id="23c8e-153">Gets the number of sheet views in this worksheet.</span></span>|
||[<span data-ttu-id="23c8e-154">getItem(key: string)</span><span class="sxs-lookup"><span data-stu-id="23c8e-154">getItem(key: string)</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|<span data-ttu-id="23c8e-155">使用工作表视图的名称获取工作表视图。</span><span class="sxs-lookup"><span data-stu-id="23c8e-155">Gets a sheet view using its name.</span></span>|
||[<span data-ttu-id="23c8e-156">getItemAt(index: number)</span><span class="sxs-lookup"><span data-stu-id="23c8e-156">getItemAt(index: number)</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|<span data-ttu-id="23c8e-157">按工作表视图在集合中的索引获取工作表视图。</span><span class="sxs-lookup"><span data-stu-id="23c8e-157">Gets a sheet view by its index in the collection.</span></span>|
||[<span data-ttu-id="23c8e-158">items</span><span class="sxs-lookup"><span data-stu-id="23c8e-158">items</span></span>](/javascript/api/excel/excel.namedsheetviewcollection#items)|<span data-ttu-id="23c8e-159">获取此集合中已加载的子项。</span><span class="sxs-lookup"><span data-stu-id="23c8e-159">Gets the loaded child items in this collection.</span></span>|
|[<span data-ttu-id="23c8e-160">Worksheet</span><span class="sxs-lookup"><span data-stu-id="23c8e-160">Worksheet</span></span>](/javascript/api/excel/excel.worksheet)|[<span data-ttu-id="23c8e-161">namedSheetViews</span><span class="sxs-lookup"><span data-stu-id="23c8e-161">namedSheetViews</span></span>](/javascript/api/excel/excel.worksheet#namedsheetviews)|<span data-ttu-id="23c8e-162">返回工作表中呈现的工作表视图的集合。</span><span class="sxs-lookup"><span data-stu-id="23c8e-162">Returns a collection of sheet views that are present in the worksheet.</span></span>|

## <a name="see-also"></a><span data-ttu-id="23c8e-163">另请参阅</span><span class="sxs-lookup"><span data-stu-id="23c8e-163">See also</span></span>

- [<span data-ttu-id="23c8e-164">Excel JavaScript API 参考文档</span><span class="sxs-lookup"><span data-stu-id="23c8e-164">Excel JavaScript API Reference Documentation</span></span>](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [<span data-ttu-id="23c8e-165">Excel JavaScript 预览 API</span><span class="sxs-lookup"><span data-stu-id="23c8e-165">Excel JavaScript preview APIs</span></span>](excel-preview-apis.md)
- [<span data-ttu-id="23c8e-166">Excel JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="23c8e-166">Excel JavaScript API requirement sets</span></span>](excel-api-requirement-sets.md)

---
title: Excel JavaScript API 高级编程概念
description: 了解 Excel 加载项如何通过使用 Office JavaScript API 对象模型与 Excel 中的对象进行交互。
ms.date: 07/01/2020
localization_priority: Priority
ms.openlocfilehash: 81602f48231f20b50a454134bc789dfdee2bbc12
ms.sourcegitcommit: 4f2f1c0a8ee777a43bb28efa226684261f4c4b9f
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/08/2020
ms.locfileid: "45081394"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="1a31e-103">Excel JavaScript API 高级编程概念</span><span class="sxs-lookup"><span data-stu-id="1a31e-103">Advanced programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="1a31e-104">本文构建于 [Excel JavaScript API 基本编程概念](excel-add-ins-core-concepts.md)中的信息之上，介绍了生成适用于 Excel 2016 或更高版本的复杂加载项所必需的一些更高级的概念。</span><span class="sxs-lookup"><span data-stu-id="1a31e-104">This article builds upon the information in [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md) to describe some of the more advanced concepts that are essential to building complex add-ins for Excel 2016 or later.</span></span>

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="1a31e-105">适用于 Excel 的 Office.js API</span><span class="sxs-lookup"><span data-stu-id="1a31e-105">Office.js APIs for Excel</span></span>

<span data-ttu-id="1a31e-106">Excel 加载项通过使用适 Office JavaScript API 与 Excel 中的对象进行交互，JavaScript API包括两个 JavaScript 对象模型：</span><span class="sxs-lookup"><span data-stu-id="1a31e-106">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="1a31e-107">**Excel JavaScript API**：[Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) 随 Office 2016 一起引入，提供了强类型的对象，可用于访问工作表、区域、表格、图表等。</span><span class="sxs-lookup"><span data-stu-id="1a31e-107">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="1a31e-108">**通用 API**：[通用 API](/javascript/api/office) 随 Office 2013 引入，可用于访问多种类型的 Office 应用程序中常见的 UI、对话框和客户端设置等功能。</span><span class="sxs-lookup"><span data-stu-id="1a31e-108">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="1a31e-109">你可能会使用 Excel JavaScript API 开发面向 Excel 2016 或更高版本的加载项中的大部分功能，同时还可以使用通用 API 中的对象。</span><span class="sxs-lookup"><span data-stu-id="1a31e-109">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API.</span></span> <span data-ttu-id="1a31e-110">例如：</span><span class="sxs-lookup"><span data-stu-id="1a31e-110">For example:</span></span>

- <span data-ttu-id="1a31e-111">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API.</span><span class="sxs-lookup"><span data-stu-id="1a31e-111">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="1a31e-112">It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span><span class="sxs-lookup"><span data-stu-id="1a31e-112">It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="1a31e-113">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span><span class="sxs-lookup"><span data-stu-id="1a31e-113">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span>

- <span data-ttu-id="1a31e-114">[Document](/javascript/api/office/office.document)：`Document` 对象提供 `getFileAsync()` 方法，用于下载运行加载项的 Excel 文件。</span><span class="sxs-lookup"><span data-stu-id="1a31e-114">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span>

<span data-ttu-id="1a31e-115">下图说明了可能使用 Excel JavaScript API 或公共 API 的情况。</span><span class="sxs-lookup"><span data-stu-id="1a31e-115">The following image illustrates when you might use the Excel JavaScript API or the Common APIs.</span></span>

![Excel JS API 和公共 API 之间差异的图像](../images/excel-js-api-common-api.png)

## <a name="requirement-sets"></a><span data-ttu-id="1a31e-117">要求集</span><span class="sxs-lookup"><span data-stu-id="1a31e-117">Requirement sets</span></span>

<span data-ttu-id="1a31e-118">Requirement sets are named groups of API members.</span><span class="sxs-lookup"><span data-stu-id="1a31e-118">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="1a31e-119">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs.</span><span class="sxs-lookup"><span data-stu-id="1a31e-119">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs.</span></span> <span data-ttu-id="1a31e-120">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="1a31e-120">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="1a31e-121">在运行时检查要求集支持</span><span class="sxs-lookup"><span data-stu-id="1a31e-121">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="1a31e-122">以下代码示例显示如何确定运行加载项的主机应用程序是否支持指定的 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="1a31e-122">The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="1a31e-123">在清单中定义要求集支持</span><span class="sxs-lookup"><span data-stu-id="1a31e-123">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="1a31e-124">可以在加载项清单中使用[要求元素](../reference/manifest/requirements.md)指定加载项要求激活的最小要求集和/或 API 方法。</span><span class="sxs-lookup"><span data-stu-id="1a31e-124">You can use the [Requirements element](../reference/manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="1a31e-125">如果 Office 主机或平台不支持清单的 `Requirements` 元素中指定的要求集或 API 方法，该加载项不会在该主机或平台中运行，而且不会显示在“我的加载项”\*\*\*\* 中显示的加载项列表中。</span><span class="sxs-lookup"><span data-stu-id="1a31e-125">If the Office host or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span></span>

<span data-ttu-id="1a31e-126">以下代码示例显示加载项清单中的 `Requirements` 元素，该元素指定应在支持 ExcelApi 要求集版本 1.3 或更高版本的所有 Office 主机应用程序中加载该加载项。</span><span class="sxs-lookup"><span data-stu-id="1a31e-126">The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="1a31e-127">为了让加载项适用于 Office 主机的所有平台（如 Excel 网页版、Windows 版 Excel 和 iPad 版 Excel），建议在运行时检查是否有要求支持，而不是在清单中定义要求集支持。</span><span class="sxs-lookup"><span data-stu-id="1a31e-127">To make your add-in available on all platforms of an Office host, such as Excel on the web, Windows, and iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="1a31e-128">Office.js 通用 API 的要求集</span><span class="sxs-lookup"><span data-stu-id="1a31e-128">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="1a31e-129">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="1a31e-129">For information about Common API requirement sets, see [Office Common API requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>

## <a name="loading-the-properties-of-an-object"></a><span data-ttu-id="1a31e-130">加载对象的属性</span><span class="sxs-lookup"><span data-stu-id="1a31e-130">Loading the properties of an object</span></span>

<span data-ttu-id="1a31e-131">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs.</span><span class="sxs-lookup"><span data-stu-id="1a31e-131">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs.</span></span> <span data-ttu-id="1a31e-132">The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span><span class="sxs-lookup"><span data-stu-id="1a31e-132">The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span></span>

### <a name="method-details"></a><span data-ttu-id="1a31e-133">方法的详细信息</span><span class="sxs-lookup"><span data-stu-id="1a31e-133">Method details</span></span>

#### `load(propertyNames?: string | string[])`

<span data-ttu-id="1a31e-134">将命令加入队列以加载对象的指定属性。</span><span class="sxs-lookup"><span data-stu-id="1a31e-134">Queues up a command to load the specified properties of the object.</span></span> <span data-ttu-id="1a31e-135">阅读属性前必须先调用 `context.sync()`。</span><span class="sxs-lookup"><span data-stu-id="1a31e-135">You must call `context.sync()` before reading the properties.</span></span>

#### <a name="syntax"></a><span data-ttu-id="1a31e-136">语法</span><span class="sxs-lookup"><span data-stu-id="1a31e-136">Syntax</span></span>

```js
object.load(param);
```

#### <a name="parameters"></a><span data-ttu-id="1a31e-137">参数</span><span class="sxs-lookup"><span data-stu-id="1a31e-137">Parameters</span></span>

|<span data-ttu-id="1a31e-138">**参数**</span><span class="sxs-lookup"><span data-stu-id="1a31e-138">**Parameter**</span></span>|<span data-ttu-id="1a31e-139">**类型**</span><span class="sxs-lookup"><span data-stu-id="1a31e-139">**Type**</span></span>|<span data-ttu-id="1a31e-140">**说明**</span><span class="sxs-lookup"><span data-stu-id="1a31e-140">**Description**</span></span>|
|:------------|:-------|:----------|
|`propertyNames`|<span data-ttu-id="1a31e-141">object</span><span class="sxs-lookup"><span data-stu-id="1a31e-141">object</span></span>|<span data-ttu-id="1a31e-142">可选。</span><span class="sxs-lookup"><span data-stu-id="1a31e-142">Optional.</span></span> <span data-ttu-id="1a31e-143">接受用逗号分隔的字符串或数组形式的属性名称。</span><span class="sxs-lookup"><span data-stu-id="1a31e-143">Accepts property names as comma-delimited string or an array.</span></span>|

#### <a name="returns"></a><span data-ttu-id="1a31e-144">返回</span><span class="sxs-lookup"><span data-stu-id="1a31e-144">Returns</span></span>

<span data-ttu-id="1a31e-145">void</span><span class="sxs-lookup"><span data-stu-id="1a31e-145">void</span></span>

#### <a name="example"></a><span data-ttu-id="1a31e-146">示例</span><span class="sxs-lookup"><span data-stu-id="1a31e-146">Example</span></span>

<span data-ttu-id="1a31e-147">The following code sample sets the properties of one Excel range by copying the properties of another range.</span><span class="sxs-lookup"><span data-stu-id="1a31e-147">The following code sample sets the properties of one Excel range by copying the properties of another range.</span></span> <span data-ttu-id="1a31e-148">Note that the source object must be loaded first, before its property values can be accessed and written to the target range.</span><span class="sxs-lookup"><span data-stu-id="1a31e-148">Note that the source object must be loaded first, before its property values can be accessed and written to the target range.</span></span> <span data-ttu-id="1a31e-149">This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span><span class="sxs-lookup"><span data-stu-id="1a31e-149">This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span></span>

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var sourceRange = sheet.getRange("B2:E2");
    sourceRange.load("format/fill/color, format/font/name, format/font/color");

    return ctx.sync()
        .then(function () {
            var targetRange = sheet.getRange("B7:E7");
            targetRange.set(sourceRange);
            targetRange.format.autofitColumns();

            return ctx.sync();
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="load-option-properties"></a><span data-ttu-id="1a31e-150">加载选项属性</span><span class="sxs-lookup"><span data-stu-id="1a31e-150">Load option properties</span></span>

<span data-ttu-id="1a31e-151">作为调用 `load()` 方法时传递逗号分隔的字符串或数组的替代方法，可以传递一个包含以下属性的对象。</span><span class="sxs-lookup"><span data-stu-id="1a31e-151">As an alternative to passing a comma-delimited string or array when you call the `load()` method, you can pass an object that contains the following properties.</span></span>

|<span data-ttu-id="1a31e-152">**属性**</span><span class="sxs-lookup"><span data-stu-id="1a31e-152">**Property**</span></span>|<span data-ttu-id="1a31e-153">**类型**</span><span class="sxs-lookup"><span data-stu-id="1a31e-153">**Type**</span></span>|<span data-ttu-id="1a31e-154">**说明**</span><span class="sxs-lookup"><span data-stu-id="1a31e-154">**Description**</span></span>|
|:-----------|:-------|:----------|
|`select`|<span data-ttu-id="1a31e-155">object</span><span class="sxs-lookup"><span data-stu-id="1a31e-155">object</span></span>|<span data-ttu-id="1a31e-156">Contains a comma-delimited list or an array of scalar property names.</span><span class="sxs-lookup"><span data-stu-id="1a31e-156">Contains a comma-delimited list or an array of scalar property names.</span></span> <span data-ttu-id="1a31e-157">Optional.</span><span class="sxs-lookup"><span data-stu-id="1a31e-157">Optional.</span></span>|
|`expand`|<span data-ttu-id="1a31e-158">object</span><span class="sxs-lookup"><span data-stu-id="1a31e-158">object</span></span>|<span data-ttu-id="1a31e-159">Contains a comma-delimited list or an array of navigational property names.</span><span class="sxs-lookup"><span data-stu-id="1a31e-159">Contains a comma-delimited list or an array of navigational property names.</span></span> <span data-ttu-id="1a31e-160">Optional.</span><span class="sxs-lookup"><span data-stu-id="1a31e-160">Optional.</span></span>|
|`top`|<span data-ttu-id="1a31e-161">int</span><span class="sxs-lookup"><span data-stu-id="1a31e-161">int</span></span>| <span data-ttu-id="1a31e-162">Specifies the maximum number of collection items that can be included in the result.</span><span class="sxs-lookup"><span data-stu-id="1a31e-162">Specifies the maximum number of collection items that can be included in the result.</span></span> <span data-ttu-id="1a31e-163">Optional.</span><span class="sxs-lookup"><span data-stu-id="1a31e-163">Optional.</span></span> <span data-ttu-id="1a31e-164">You can only use this option when you use the object notation option.</span><span class="sxs-lookup"><span data-stu-id="1a31e-164">You can only use this option when you use the object notation option.</span></span>|
|`skip`|<span data-ttu-id="1a31e-165">int</span><span class="sxs-lookup"><span data-stu-id="1a31e-165">int</span></span>|<span data-ttu-id="1a31e-166">Specify the number of items in the collection that are to be skipped and not included in the result.</span><span class="sxs-lookup"><span data-stu-id="1a31e-166">Specify the number of items in the collection that are to be skipped and not included in the result.</span></span> <span data-ttu-id="1a31e-167">If `top` is specified, the result set will start after skipping the specified number of items.</span><span class="sxs-lookup"><span data-stu-id="1a31e-167">If `top` is specified, the result set will start after skipping the specified number of items.</span></span> <span data-ttu-id="1a31e-168">Optional.</span><span class="sxs-lookup"><span data-stu-id="1a31e-168">Optional.</span></span> <span data-ttu-id="1a31e-169">You can only use this option when you use the object notation option.</span><span class="sxs-lookup"><span data-stu-id="1a31e-169">You can only use this option when you use the object notation option.</span></span>|

<span data-ttu-id="1a31e-170">以下代码示例通过为集合中的每个工作表的所用区域选择 `name` 属性和 `address` 来加载工作表集合。</span><span class="sxs-lookup"><span data-stu-id="1a31e-170">The following code sample loads a worksheet collection by selecting the `name` property and the `address` of the used range for each worksheet in the collection.</span></span> <span data-ttu-id="1a31e-171">它还指定只能加载集合中的前五个工作表。</span><span class="sxs-lookup"><span data-stu-id="1a31e-171">It also specifies that only the top five worksheets in the collection should be loaded.</span></span> <span data-ttu-id="1a31e-172">可以通过将 `top: 10` 和 `skip: 5` 指定为属性值来处理下一组五个工作表。</span><span class="sxs-lookup"><span data-stu-id="1a31e-172">You could process the next set of five worksheets by specifying `top: 10` and `skip: 5` as attribute values.</span></span>

```js
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

### <a name="calling-load-without-parameters"></a><span data-ttu-id="1a31e-173">不带参数调用 `load`</span><span class="sxs-lookup"><span data-stu-id="1a31e-173">Calling `load` without parameters</span></span>

<span data-ttu-id="1a31e-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded.</span><span class="sxs-lookup"><span data-stu-id="1a31e-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded.</span></span> <span data-ttu-id="1a31e-175">To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span><span class="sxs-lookup"><span data-stu-id="1a31e-175">To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1a31e-176">无参数 `load` 语句返回的数据量可能超过该服务的大小限制。</span><span class="sxs-lookup"><span data-stu-id="1a31e-176">The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service.</span></span> <span data-ttu-id="1a31e-177">为了降低较旧加载项的风险，`load` 不会在明确请求它们之前返回某些属性。</span><span class="sxs-lookup"><span data-stu-id="1a31e-177">To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them.</span></span> <span data-ttu-id="1a31e-178">此类加载操作中排除了以下属性：</span><span class="sxs-lookup"><span data-stu-id="1a31e-178">The following properties are excluded from such load operations:</span></span>
>
> * `Excel.Range.numberFormatCategories`

## <a name="scalar-and-navigation-properties"></a><span data-ttu-id="1a31e-179">标量和导航属性</span><span class="sxs-lookup"><span data-stu-id="1a31e-179">Scalar and navigation properties</span></span>

<span data-ttu-id="1a31e-180">属性分为两种类别：**标量**和**导航**。</span><span class="sxs-lookup"><span data-stu-id="1a31e-180">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="1a31e-181">标量属性是可分配的类型，如字符串、整数和 JSON 结构。</span><span class="sxs-lookup"><span data-stu-id="1a31e-181">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="1a31e-182">导航属性是只读对象和已分配字段的对象的集合，而不是直接分配属性。</span><span class="sxs-lookup"><span data-stu-id="1a31e-182">Navigation properties are readonly objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="1a31e-183">例如，[Worksheet](/javascript/api/excel/excel.worksheet) 对象上的 `name`和 `position` 成员是标量属性，而 `protection` 和 `tables` 是导航属性。</span><span class="sxs-lookup"><span data-stu-id="1a31e-183">For example, `name` and `position` members on the [Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.</span></span> <span data-ttu-id="1a31e-184">[DataValidation](/javascript/api/excel/excel.datavalidation) 对象上的 `prompt` 是必须使用 JSON 对象 (`dv.prompt = { title: "MyPrompt"}`) 设置的标量属性的示例，而不是设置子属性 (`dv.prompt.title = "MyPrompt" // will not set the title`)。</span><span class="sxs-lookup"><span data-stu-id="1a31e-184">`prompt` on the [DataValidation](/javascript/api/excel/excel.datavalidation) object is an example of a scalar property that must be set using a JSON object (`dv.prompt = { title: "MyPrompt"}`), instead of setting the sub-properties (`dv.prompt.title = "MyPrompt" // will not set the title`).</span></span>

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a><span data-ttu-id="1a31e-185">使用 `object.load()` 的标量属性和导航属性</span><span class="sxs-lookup"><span data-stu-id="1a31e-185">Scalar properties and navigation properties with `object.load()`</span></span>

<span data-ttu-id="1a31e-186">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded.</span><span class="sxs-lookup"><span data-stu-id="1a31e-186">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded.</span></span> <span data-ttu-id="1a31e-187">Additionally, navigation properties cannot be loaded directly.</span><span class="sxs-lookup"><span data-stu-id="1a31e-187">Additionally, navigation properties cannot be loaded directly.</span></span> <span data-ttu-id="1a31e-188">Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property.</span><span class="sxs-lookup"><span data-stu-id="1a31e-188">Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property.</span></span> <span data-ttu-id="1a31e-189">For example, to load the font name for a range, you must specify the `format` and `font` navigation properties as the path to the `name` property:</span><span class="sxs-lookup"><span data-stu-id="1a31e-189">For example, to load the font name for a range, you must specify the `format` and `font` navigation properties as the path to the `name` property:</span></span>

```js
someRange.load("format/font/name")
```

> [!NOTE]
> <span data-ttu-id="1a31e-190">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path.</span><span class="sxs-lookup"><span data-stu-id="1a31e-190">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="1a31e-191">For example, you could set the font size for a range by using `someRange.format.font.size = 10;`.</span><span class="sxs-lookup"><span data-stu-id="1a31e-191">For example, you could set the font size for a range by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="1a31e-192">You do not need to load the property before you set it.</span><span class="sxs-lookup"><span data-stu-id="1a31e-192">You do not need to load the property before you set it.</span></span> 

## <a name="setting-properties-of-an-object"></a><span data-ttu-id="1a31e-193">设置对象的属性</span><span class="sxs-lookup"><span data-stu-id="1a31e-193">Setting properties of an object</span></span>

<span data-ttu-id="1a31e-194">Setting properties on an object with nested navigation properties can be cumbersome.</span><span class="sxs-lookup"><span data-stu-id="1a31e-194">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="1a31e-195">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API.</span><span class="sxs-lookup"><span data-stu-id="1a31e-195">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API.</span></span> <span data-ttu-id="1a31e-196">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span><span class="sxs-lookup"><span data-stu-id="1a31e-196">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

> [!NOTE]
> <span data-ttu-id="1a31e-197">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API.</span><span class="sxs-lookup"><span data-stu-id="1a31e-197">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API.</span></span> <span data-ttu-id="1a31e-198">The common (shared) APIs do not support this method.</span><span class="sxs-lookup"><span data-stu-id="1a31e-198">The common (shared) APIs do not support this method.</span></span> 

### <a name="set-properties-object-options-object"></a><span data-ttu-id="1a31e-199">set (properties: object, options: object)</span><span class="sxs-lookup"><span data-stu-id="1a31e-199">set (properties: object, options: object)</span></span>

<span data-ttu-id="1a31e-200">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object.</span><span class="sxs-lookup"><span data-stu-id="1a31e-200">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object.</span></span> <span data-ttu-id="1a31e-201">If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span><span class="sxs-lookup"><span data-stu-id="1a31e-201">If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span></span>

#### <a name="syntax"></a><span data-ttu-id="1a31e-202">语法</span><span class="sxs-lookup"><span data-stu-id="1a31e-202">Syntax</span></span>

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a><span data-ttu-id="1a31e-203">参数</span><span class="sxs-lookup"><span data-stu-id="1a31e-203">Parameters</span></span>

|<span data-ttu-id="1a31e-204">**参数**</span><span class="sxs-lookup"><span data-stu-id="1a31e-204">**Parameter**</span></span>|<span data-ttu-id="1a31e-205">**类型**</span><span class="sxs-lookup"><span data-stu-id="1a31e-205">**Type**</span></span>|<span data-ttu-id="1a31e-206">**说明**</span><span class="sxs-lookup"><span data-stu-id="1a31e-206">**Description**</span></span>|
|:------------|:--------|:----------|
|`properties`|<span data-ttu-id="1a31e-207">object</span><span class="sxs-lookup"><span data-stu-id="1a31e-207">object</span></span>|<span data-ttu-id="1a31e-208">与在其上调用方法的对象相同的 Office.js 类型的对象，或属性名称及类型反映在其上调用方法的对象结构的 JavaScript 对象。</span><span class="sxs-lookup"><span data-stu-id="1a31e-208">Either an object of the same Office.js type of the object on which the method is called, or a JavaScript object with property names and types that mirror the structure of the object on which the method is called.</span></span>|
|`options`|<span data-ttu-id="1a31e-209">object</span><span class="sxs-lookup"><span data-stu-id="1a31e-209">object</span></span>|<span data-ttu-id="1a31e-210">Optional.</span><span class="sxs-lookup"><span data-stu-id="1a31e-210">Optional.</span></span> <span data-ttu-id="1a31e-211">Can only be passed when the first parameter is a JavaScript object.</span><span class="sxs-lookup"><span data-stu-id="1a31e-211">Can only be passed when the first parameter is a JavaScript object.</span></span> <span data-ttu-id="1a31e-212">The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span><span class="sxs-lookup"><span data-stu-id="1a31e-212">The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span></span>|

#### <a name="returns"></a><span data-ttu-id="1a31e-213">返回</span><span class="sxs-lookup"><span data-stu-id="1a31e-213">Returns</span></span>

<span data-ttu-id="1a31e-214">void</span><span class="sxs-lookup"><span data-stu-id="1a31e-214">void</span></span>

#### <a name="example"></a><span data-ttu-id="1a31e-215">示例</span><span class="sxs-lookup"><span data-stu-id="1a31e-215">Example</span></span>

<span data-ttu-id="1a31e-216">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object.</span><span class="sxs-lookup"><span data-stu-id="1a31e-216">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object.</span></span> <span data-ttu-id="1a31e-217">This example assumes that there is data in range **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="1a31e-217">This example assumes that there is data in range **B2:E2**.</span></span>

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="42ornullobject-methods"></a><span data-ttu-id="1a31e-218">&#42;OrNullObject 方法</span><span class="sxs-lookup"><span data-stu-id="1a31e-218">&#42;OrNullObject methods</span></span>

<span data-ttu-id="1a31e-219">Many Excel JavaScript API methods will return an exception when the condition of the API is not met.</span><span class="sxs-lookup"><span data-stu-id="1a31e-219">Many Excel JavaScript API methods will return an exception when the condition of the API is not met.</span></span> <span data-ttu-id="1a31e-220">For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span><span class="sxs-lookup"><span data-stu-id="1a31e-220">For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span></span> 

<span data-ttu-id="1a31e-221">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API.</span><span class="sxs-lookup"><span data-stu-id="1a31e-221">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API.</span></span> <span data-ttu-id="1a31e-222">An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist.</span><span class="sxs-lookup"><span data-stu-id="1a31e-222">An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist.</span></span> <span data-ttu-id="1a31e-223">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection.</span><span class="sxs-lookup"><span data-stu-id="1a31e-223">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection.</span></span> <span data-ttu-id="1a31e-224">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object.</span><span class="sxs-lookup"><span data-stu-id="1a31e-224">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object.</span></span> <span data-ttu-id="1a31e-225">The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span><span class="sxs-lookup"><span data-stu-id="1a31e-225">The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span></span>

<span data-ttu-id="1a31e-226">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method.</span><span class="sxs-lookup"><span data-stu-id="1a31e-226">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="1a31e-227">If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span><span class="sxs-lookup"><span data-stu-id="1a31e-227">If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span></span>

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
  .then(function() {
    if (dataSheet.isNullObject) {
        // Create the sheet
    }

    dataSheet.position = 1;
    //...
  })
```

## <a name="see-also"></a><span data-ttu-id="1a31e-228">另请参阅</span><span class="sxs-lookup"><span data-stu-id="1a31e-228">See also</span></span>

* [<span data-ttu-id="1a31e-229">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="1a31e-229">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="1a31e-230">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="1a31e-230">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="1a31e-231">Excel JavaScript API 性能优化</span><span class="sxs-lookup"><span data-stu-id="1a31e-231">Excel JavaScript API performance optimization</span></span>](performance.md)
* [<span data-ttu-id="1a31e-232">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="1a31e-232">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
